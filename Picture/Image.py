# /usr/bin/env python
# -*- coding:utf-8 -*-

# https://stackoverflow.com/questions/7829499/using-hashlib-to-compute-md5-digest-of-a-file-in-python-3
# https://www.cnblogs.com/ivictor/p/4377609.html
# https://zhuanlan.zhihu.com/p/67959761
# https://www.huaweicloud.com/articles/b4f513903329669a822a4414d67b19a0.html
# https://docs.python.org/zh-cn/3/library/multiprocessing.html?highlight=queue
# https://pypi.org/project/ExifRead/
# https://www.jianshu.com/p/3b61923efdf1
# https://blog.csdn.net/weixin_43745169/article/details/100988915
import exifread
import os
import shutil
import hashlib
from functools import partial
from multiprocessing import Pool, Manager
import glob
from hachoir import parser, metadata
from geopy.geocoders import Nominatim
from configparser import ConfigParser

# 自定义程序运行参数
mime_types = [
    'jpeg', 'png', 'jpg', 'webp', 'gif', 'bmp', 'mp4', 'mov', 'm4v', 'wmv',
    'avi', 'rm', 'rmvb'
]
root = os.getcwd()
dist_root = os.path.join(root, 'dist')
no_create_time_root = os.path.join(root, '无拍摄时间')
log_file = os.path.join(root, '重复文件记录.txt')
config_path = os.path.join(root, 'config.ini')
###################################################################

geolocator = Nominatim(user_agent="imagepy")


def parse_config():
    global mime_types, root, dist_root, no_create_time_root, log_file
    if not os.path.exists(config_path):
        return
    cfg = ConfigParser()
    try:
        cfg.read('./config.ini')
    except:
        return
    try:
        mime_types = cfg['common']['mime_types'].split(',')
    except:
        pass
    try:
        root = cfg['common']['root']
    except:
        pass
    try:
        dist_root = cfg['common']['dist']
    except:
        pass
    try:
        no_create_time_root = cfg['common']['fallback_dist']
    except:
        pass
    try:
        log_file = cfg.get('common', 'log_file')
    except:
        pass


def ensure_path():
    global mime_types
    mime_types.extend(list(map(lambda s: s.upper(), mime_types)))
    if not os.path.exists(no_create_time_root):
        os.mkdir(no_create_time_root)
    if not os.path.exists(dist_root):
        os.mkdir(dist_root)


def get_images_path(media_type, images, lock):
    pattern = os.path.join(root, '**/*.' + media_type)
    images_part = glob.glob(pattern, recursive=True)
    lock.acquire()
    images.extend(images_part)
    lock.release()


def check_md5(path, f):
    d = hashlib.md5()
    for buf in iter(partial(f.read, 8 * 1024), b''):
        d.update(buf)
    return d


def output_duplicated_files(duplicated_files):
    if len(duplicated_files.values()) > 0:
        with open(log_file, 'w') as f:
            for item in duplicated_files.values():
                f.write(item.replace(',', '\n') + '\n\n')
        print('存在重复文件，详情查看：重复文件记录.txt')


def geo_parse(latitude, longitude):
    if not isinstance(latitude, str) or latitude == '' or not isinstance(
            longitude, str) or longitude == '':
        return ''
    latitude_longitude = latitude + ',' + longitude
    try:
        position = geolocator.reverse(latitude_longitude)
        address = position.address
    except:
        return ''
    address = address.split(', ')
    address.reverse()
    address.pop(1)
    return ''.join(address)


def get_media_metas(filePath):
    parserFile = parser.createParser(filePath)  #解析文件
    if not parserFile:
        return ''
    try:
        metadataDecode = metadata.extractMetadata(parserFile)  # 获取文件的metadata
    except ValueError:
        return ''
    if not metadataDecode:
        return ''
    metaInfos = metadataDecode.exportPlaintext(
        line_prefix="")  # 将文件的metadata转换为list,且将前缀设置为空
    creation_date = ''
    date_time_original = ''
    latitude = ''
    longitude = ''
    for meta in metaInfos:
        #如果字符串在列表中,则提取数字部分,即为文件创建时间
        if 'Creation date' in meta:
            creation_date = meta.replace('Creation date: ',
                                         '').replace(' ', '_')
        elif 'Date-time original' in meta:
            date_time_original = meta.replace('Date-time original: ',
                                              '').replace(' ', '_')
        elif 'Latitude' in meta:
            latitude = meta.replace('Latitude: ', '')
        elif 'Longitude' in meta:
            longitude = meta.replace('Longitude: ', '')
    geo_address = geo_parse(latitude, longitude)
    return ((date_time_original or creation_date), geo_address)


def move_file(path, record_files, duplicated_files, lock):
    with open(path, 'rb') as f:
        md5 = check_md5(path, f)
    unique_key = (os.path.getsize(path), md5.hexdigest())

    lock.acquire()
    if unique_key in record_files:
        if unique_key in duplicated_files:
            duplicated_files[
                unique_key] = duplicated_files[unique_key] + ',' + path
        else:
            duplicated_files[
                unique_key] = record_files[unique_key] + ',' + path
        lock.release()
        return
    record_files[unique_key] = '**-' + unique_key[1][8:24] + os.path.splitext(
        path)[1]
    lock.release()
    (create_time, geo_address) = get_media_metas(path)
    if geo_address != '':
        geo_address = '_' + geo_address
    dist_dir = no_create_time_root
    filename_prefix = unique_key[1][8:24]
    if create_time != '':
        year_month = create_time[:7]
        year = year_month[:4]
        # 如果年份小于2000年明显不对
        if int(year) > 2000:
            filename_prefix = create_time + '_' + unique_key[1][8:24]
            dist_dir = os.path.join(dist_root, year + '/' + year_month)
            if not os.path.exists(dist_dir):
                os.makedirs(dist_dir)

    dist_path = os.path.join(
        dist_dir, filename_prefix + geo_address + os.path.splitext(path)[1])
    if not os.path.exists(dist_path):
        shutil.move(path, dist_path)
    else:
        # 可能之前该目录已经存在被处理的文件
        lock.acquire()
        if unique_key in duplicated_files:
            duplicated_files[
                unique_key] = duplicated_files[unique_key] + ',' + path
        else:
            duplicated_files[
                unique_key] = record_files[unique_key] + ',' + path
        lock.release()


parse_config()
ensure_path()

if __name__ == "__main__":
    print('开始处理...')
    cpu_count = os.cpu_count()
    images = []
    with Manager() as manager1:
        images_proxy = manager1.list()
        lock1 = manager1.Lock()
        pl1 = Pool(cpu_count - 1 if cpu_count - 1 > 1 else 1)
        for t in mime_types:
            pl1.apply_async(get_images_path, args=(t, images_proxy, lock1))
        pl1.close()
        pl1.join()
        images.extend(images_proxy)

    with Manager() as manager2:
        record_files = manager2.dict()  #保存已记录的文件信息
        duplicated_files = manager2.dict()  #保存重复的文件信息
        lock2 = manager2.Lock()

        pl2 = Pool(cpu_count - 1 if cpu_count - 1 > 1 else 1)
        for image_path in images:
            pl2.apply_async(move_file,
                            args=(image_path, record_files, duplicated_files,
                                  lock2))
        pl2.close()
        pl2.join()
        output_duplicated_files(duplicated_files)

    print('处理完成啦!')
import datetime
import os
import exifread

path = r'./'

photo_suffixes = ['.jpg', '.png']
exif_keys = ['Image DateTime', 'EXIF DateTimeOriginal']

files = os.listdir(path)
count, files_number = 0, len(files)
for item in files:
    print('processing %d/%d:' % (count, files_number), end=' ')
    count += 1

    src_path = os.path.join(os.path.abspath(path), item)
    name, suffix = os.path.splitext(item)
    if suffix.lower() not in photo_suffixes:
        print('%s is not image.' % src_path)
        continue

    with open(src_path, 'rb') as image:
        exif_tags = exifread.process_file(image)

    image_time = None
    for key in exif_keys:
        if key not in exif_tags:
            continue
        image_time = datetime.datetime.strptime(exif_tags[key].values,
                                                "%Y:%m:%d %H:%M:%S")

    if image_time is None:
        print('%s don\'t have exif info.' % src_path)
        continue

    format_time = image_time.strftime('%Y%m%d_%H%M%S')
    new_path = os.path.join(os.path.abspath(path), format_time + suffix)

    if new_path == src_path:
        print('%s don\'t need rename.' % src_path)
        continue

    deduplicate_count = 2
    while os.path.isfile(new_path):
        new_path = os.path.join(
            os.path.abspath(path),
            '%s_%d%s' % (format_time, deduplicate_count, suffix))
        deduplicate_count += 1
    os.rename(src_path, new_path)

    print('rename %s to %s ...' % (src_path, new_path))

os.system('pause')

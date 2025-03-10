import os
import shutil

all_test_folder = 'all_test_results'

if not os.path.exists(all_test_folder):
    os.mkdir(all_test_folder)


folders = ('total_files','otherResult')
count = 0
def get_inner(path_folder):
    global count
    inners_folder = os.listdir(path_folder)
    for inner in inners_folder:
        path_inner = os.path.join(path_folder,inner)
        if 'xlsx' in inner:
            count += 1
            new_path = os.path.join(all_test_folder,str(count)+inner)
            shutil.copy(path_inner,new_path)
        if os.path.isdir(path_inner):
            get_inner(path_inner)

for folder in folders:
    get_inner(folder)


# # Пути к исходному и целевому файлу
# source_file = "путь/к/исходному/файлу.txt"
# destination_file = "путь/к/целевому/файлу.txt"

# # Копирование файла
# shutil.copy(source_file, destination_file)
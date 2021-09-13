from openpyxl import Workbook, load_workbook
import os
from pathlib import Path


# Task 1: Write code to rename the files based on our dictionary
## Task 1.1: Pull video IDs and Titles from excel files to your script
path = '/home/lee/Desktop/Mini-Project_2_Renames_and_Organize_files_directories_bot/organizeVideos/database.xlsx'
workbook = load_workbook(path)
video_test_sheet_index = 1
video_test_sheet = workbook.get_sheet_names()[video_test_sheet_index]
worksheet = workbook.get_sheet_by_name(video_test_sheet)


video_name_sheet_index = 0
video_name_sheet = workbook.get_sheet_names()[video_name_sheet_index]
worksheet = workbook.get_sheet_by_name(video_name_sheet)


# First column contain video titles
first_column = worksheet['A']
list_video_titles = []
start_row = 1
end_row = len(first_column)
for row  in range (start_row, end_row):
    video_title = first_column[row].value
    if video_title is not None:
        list_video_titles.append(video_title)
print(len(list_video_titles))


# Second column contain video id
second_column = worksheet['B']
list_video_ids = []
start_row = 1
end_row = len(second_column)
for row in range (start_row, end_row):
    video_id = second_column[row].value
    if video_id is not None:
        list_video_ids.append(int(video_id))
print(len(list_video_ids))


list_video_ids_str = []
for video_id in list_video_ids:
    # Can not compare str and float
    video_id = int(video_id)
    video_id = str(video_id)
    list_video_ids_str.append(video_id)
print (len(list_video_ids_str))


## Task 1.2:  Match the video IDs to its corresponding titles
id_dict = {}
for index in range (len(list_video_ids_str)):
    key = list_video_ids_str[index]
    value = list_video_titles[index]
    id_dict[key] = value
    # id_dict.update({list_video_ids_str[index]:list_video_titles[index]})
len(id_dict)


## Task 1.3: Renaming files from its IDs to its Titles
file_path = '/home/lee/Desktop/Mini-Project_2_Renames_and_Organize_files_directories_bot/organizeVideos/video'
os.chdir(file_path)

list_video = os.listdir()
for video in list_video:
    for id, title in id_dict.items():
        # Split name and file extension from video name
        # Ex: 3434.mp4 so video_name = "3434", file_extension = ".mp4"
        video_name, file_extension = os.path.splitext(video)
        if id == video_name:
            # f is F-string type
            name = f'{title}{file_extension}'
            print('From:', video, 'To:', name)
            os.rename(video, name)


# Task 2: Write code to remove the files to the folder of each genre
## Task 2.1 : Pull the genres data to our file
genres_sheet_index = 2
third_sheet = workbook.get_sheet_names()[genres_sheet_index]
worksheet_genre =workbook.get_sheet_by_name(third_sheet)
list_genres = []
start_row = 1
end_row = len(worksheet_genre['A'])
for row in range(start_row,end_row):
    genre = worksheet_genre['A'][row].value
    list_genres.append(genre)
print(list_genres)


## Task 2.2: Match the files with their genre
genre_dict= {}
for genre  in list_genres:
    key = genre
    # many movies can be same genre, so value is a list of movie 
    value = []
    genre_dict[key] = value


path2 ='/home/lee/Desktop/Mini-Project_2_Renames_and_Organize_files_directories_bot/organizeVideos/video'
list_video = os.listdir(path2)
for each_video in list_video:
    # key = genre, values is a list of movie video
    for key, values in genre_dict.items():
        if key in each_video:
            values.append(each_video )
def check_genre(movie_title):
    for genre in list_genres:
        if genre in  movie_title:
            return genre
    return "Others"


os.chdir(path2)
for file in list_video:
    file_from = Path(file)
    directory = check_genre(str(file))
    dirPath = Path(directory)
    if dirPath.exists() is False:
        os.mkdir(str(directory))
    file_to = dirPath.joinpath(file_from)
    os.rename(file_from,file_to)
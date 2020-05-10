# from openpyxl import load_workbook
# from dashboard.my_secrets import REGRESSION_DEVEXT_DBQA_GIS
#
#
# def get_items_from_folder(
#     gis_obj, folder, item_types=None
# ) -> list:  # folder=None returns the root folder
#     items = gis_obj.users.me.items(folder=folder)
#
#     if item_types:
#         items = [item for item in items if item.type in item_types]
#         return items
#
#     return items
#
#
# if __name__ == "__main__":
#
#     WB = load_workbook("name.xlsx")
#     WS = WB.active
#
#     FOLDERS = REGRESSION_DEVEXT_DBQA_GIS.users.me.folders
#     print(f"I have {len(FOLDERS)} folders to work on")
#
#     ALL_ITEMS = []
#     for FOLDER_INDEX, FOLDER in enumerate(FOLDERS):
#         if FOLDER["title"] == "_Trash_Can":
#             continue
#         print(f"I am working on Folder {FOLDER_INDEX} with name, {FOLDER['title']}")
#         FOLDER_ITEMS = get_items_from_folder(
#             gis_obj=REGRESSION_DEVEXT_DBQA_GIS, folder=FOLDER, item_types=["Dashboard"]
#         )
#         for ITEM in FOLDER_ITEMS:
#             try:
#                 if ITEM.get_data() is not None:
#                     ITEMS_WITH_FOLDER_AND_VERSION = (
#                         ITEM,
#                         FOLDER["title"],
#                         ITEM.get_data()["version"],
#                     )
#                 ALL_ITEMS.append(
#                     ITEMS_WITH_FOLDER_AND_VERSION
#                 )  # I think this is in the wrong spot
#             except Exception as e:
#                 print(e)
#
#     for ITEM_INDEX, row in enumerate(
#         WS.iter_rows(min_row=2, max_row=len(ALL_ITEMS) + 1)
#     ):
#         row[0].value = ALL_ITEMS[ITEM_INDEX][0].title
#         row[1].value = ALL_ITEMS[ITEM_INDEX][0].id
#         row[2].value = ALL_ITEMS[ITEM_INDEX][1]  # folder
#         row[3].value = ALL_ITEMS[ITEM_INDEX][0].type
#         row[4].value = ALL_ITEMS[ITEM_INDEX][2]
#
#     WB.save("name.xlsx")

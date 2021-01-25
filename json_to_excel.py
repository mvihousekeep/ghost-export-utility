from datetime import datetime
import json
import pandas as pd
import iso8601 as iso
from dateutil import tz
import platform


def generate_excel(file_loc, export_loc):
    # file_loc = r"C:\Users\user\PycharmProjects\MVIPostToExcel\mission-victory-india.ghost.2020-12-19-15-33-27.json"
    ist = tz.gettz("Asia/Calcutta")
    curr_dt = datetime.now().astimezone(ist).replace(tzinfo=None)
    # export_loc = r"C:\Users\user\PycharmProjects\MVIPostToExcel"
    export_name = "MVIGhostExport_" + curr_dt.strftime("%d-%m-%Y_%H-%M-%S") + ".xlsx"

    print("Starting Excel Export")
    print("Time: " + str(curr_dt))
    print("Loading Exported JSON Data File from " + file_loc)
    json_file = open(file_loc, "r", encoding="utf8")
    print("Loaded Exported JSON Data File Successfully")

    print("Creating JSON Object from File")
    json_object = json.load(json_file)
    print("Created JSON Object from File Successfully")

    print("Creating DataFrame with Columns: " + "SNo, " + "Title, " + "Slug, " + "Type, " + "Status, " +
          "Created At (IST), " + "Updated At (IST), " + "Published At (IST), " + "Author, " + "Ghost Editor Link")
    dataframe = pd.DataFrame(columns=["SNo", "Title", "Slug", "Type", "Status", "Created At (IST)", "Updated At (IST)",
                                      "Published At (IST)", "Author", "Ghost Editor Link"])
    summary = pd.DataFrame(columns=["Export Date (IST)", "Exported Records", "Input JSON Path", "Excel Export Path"])
    print("Created DataFrame Successfully")

    if json_object['db']:
        print("Found db object in JSON")
        y = 0
        for i in json_object['db']:
            y += 1
            if i['data']:
                print("Found data object in JSON[db]")
                if i['data']['posts']:
                    print("Found posts object in JSON[db][data]")

                    print("Populating DataFrame")
                    x = 0
                    author_list = i['data']['users']
                    author_map = {}
                    base_editor_url = "https://missionvictoryindia.com/ghost/#/editor/"
                    for author in author_list:
                        author_map[author['id']] = author['name']
                    for post in i['data']['posts']:
                        x += 1
                        author_name = author_map[post['author_id']]
                        dataframe.loc[x] = [x, post['title'], post['slug'], post['type'], post['status'],
                                            None if post['created_at'] is None else
                                            iso.parse_date(post['created_at']).astimezone(ist).replace(tzinfo=None),
                                            None if post['updated_at'] is None else
                                            iso.parse_date(post['updated_at']).astimezone(ist).replace(tzinfo=None),
                                            None if post['published_at'] is None else
                                            iso.parse_date(post['published_at']).astimezone(ist).replace(tzinfo=None),
                                            author_name, base_editor_url + post['type'] + "/" + post['id']]
                    summary.loc[y] = [curr_dt, x, file_loc,
                                      export_loc + ("\\" if platform.system() == 'Windows' else "/") + export_name]
                    print("Populated DataFrame SuccessFully with " + str(x) + " records")

    print("Writing DataFrame to Excel")
    writer = pd.ExcelWriter(export_loc + ("\\" if platform.system() == 'Windows' else "/") + export_name,
                            engine='xlsxwriter',
                            datetime_format='mmm d yyyy hh:mm:ss',
                            date_format='mmmm dd yyyy',
                            options={'remove_timezone': True})
    summary = summary.T
    summary.to_excel(writer, header=False, sheet_name='Export Summary', startcol=6, startrow=9)
    dataframe.to_excel(writer, index=False, sheet_name='Exported Data')
    writer.save()
    print("Excel Created Successfully")


if __name__ == '__main__':
    export_loc_test = r"C:\Users\user\PycharmProjects\MVIPostToExcel"
    file_loc_test = r"C:\Users\user\PycharmProjects\MVIPostToExcel\mission-victory-india.ghost.2021-01-19-14-07-03.json"
    generate_excel(file_loc_test, export_loc_test)

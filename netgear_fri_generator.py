import os
import netgear_queries as queries


def checkDir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def writeFRI(fri_filepath, major_equipment, manufacturing, equipment_category, equipment_type, unit_category,
             serial_number, web_defined_status, description, mobile_defined_status, remark, site_id_excel,
             barcode_flag, movement_flag, unit_of_measurement, flag_waste, req_qty_db, req_qty):
    if not os.path.exists(fri_filepath):
        text = "text|p1_1_" + serial_number + "|" + major_equipment + \
               "|@|text|p1_2_" + serial_number + "|" + manufacturing + \
               "|@|text|p1_3_" + serial_number + "|" + equipment_category + \
               "|@|text|p1_4_" + serial_number + "|" + equipment_type + \
               "|@|text|p1_5_" + serial_number + "|" + unit_category + \
               "|@|text|p1_6_" + serial_number + "|" + serial_number + \
               "|@|text|p1_7_" + serial_number + "|" + web_defined_status + \
               "|@|text|p1_8_" + serial_number + "|" + description + \
               "|@|text|p1_9_" + serial_number + "|" + mobile_defined_status + \
               "|@|text|p1_10_" + serial_number + "|" + remark + \
               "|@|text|site_id|" + site_id_excel + \
               "|@|text|barcode_flag|" + barcode_flag + \
               "|@|text|movement_flag|" + movement_flag + \
               "|@|text|unit_of_measurement|" + unit_of_measurement + \
               "|@|text|flag_waste|" + flag_waste + \
               "|@|text|req_qty_db|" + req_qty_db + \
               "|@|text|req_qty|" + req_qty + '|'

        f = open(fri_filepath, "w", encoding="utf8")
        f.write(text)
        f.close()

        fri_filename = os.path.basename(fri_filepath)
        print("\tFRI FILE CREATED:\t\t%s" % fri_filename)
    else:
        print("\tFRI FILE ALREADY EXISTS")


if __name__ == '__main__':

    #################################

    task_id = '20200406115410818'  # PLEASE INPUT TASK ID

    print("PROCESSING TASK %s" % (task_id))

    #################################

    output_path = "D:\\KERJAAN\\SCRIPTS\\SCRIPTS\\Made by AMG\\MANUAL - NETgear ATF Uploader\\OUTPUT\\GENERATED FRI"

    task_id_folder = output_path + "\\" + task_id
    checkDir(task_id_folder)

    data = queries.getFRIData(task_id)

    if data:
        print("FRI DATA ACQUIRED : %s ROW(s)" % (str(len(data))))

        for datum in data:
            major_equipment = datum[0]
            manufacturing = datum[1]
            equipment_category = datum[2]
            equipment_type = datum[3]
            unit_category = datum[4]
            serial_number = datum[5]
            web_defined_status = datum[6]
            description = datum[7]
            mobile_defined_status = datum[8]
            remark = datum[9]
            site_id_fri = datum[10]
            barcode_flag = datum[11]
            movement_flag = datum[12]
            unit_of_measurement = datum[13]
            flag_waste = datum[14]
            req_qty_db = datum[15]
            req_qty = datum[16]

            print("PROCESSING SN %s" % (serial_number))

            fri_filepath = task_id_folder + "\\" + "%s_%s_%s.fri" % (task_id, site_id_fri, serial_number)
            fri_content = (major_equipment, manufacturing, equipment_category, equipment_type, unit_category,
                           serial_number, web_defined_status, description, mobile_defined_status, remark,
                           site_id_fri, barcode_flag, movement_flag, unit_of_measurement, flag_waste,
                           req_qty_db, req_qty)
            writeFRI(fri_filepath, *fri_content)
    else:
        print("NO DATA AVAILABLE")

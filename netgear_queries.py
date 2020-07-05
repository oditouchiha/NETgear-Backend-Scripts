import os
from configparser import ConfigParser

import psycopg2

# from dito_logger import DitoLogger

# dl = DitoLogger(os.path.basename(__file__))
# logger = dl.logger


def connect_to_database_netgear():
    # ================================
    # ESTABLISH CONNECTION TO DATABASE
    # ================================

    parser = ConfigParser()
    parser.read('netgear_db_config.ini')

    try:
        # conf_section = 'db.dummy'
        conf_section = 'db.info'
        connection = psycopg2.connect(user=parser[conf_section]['user'],
                                      password=parser[conf_section]['password'],
                                      host=parser[conf_section]['host'],
                                      port=parser[conf_section]['port'],
                                      database=parser[conf_section]['database'])
        return connection

    except (Exception, psycopg2.Error) as error:
        print("[ERROR OCCURRED !!!] Error while fetching data from PostgreSQL", error)


def generateTaskID(destination=False):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT
        CASE
            WHEN (MAX(SUBSTRING(task_id, 0, 13))::BIGINT || '_atf') IS NULL THEN CONCAT(TO_CHAR(NOW(), 'YYYYMMDD'), '0001_atf')
            WHEN (MAX(SUBSTRING(task_id, 0, 13))::BIGINT || '_atf') IS NOT NULL THEN CONCAT(MAX(SUBSTRING(task_id, 0, 13))::BIGINT+1, '_atf')
        END as generated_task_id
        FROM tbl_task
        WHERE
            (SUBSTRING(task_id, 0, 9))=TO_CHAR(NOW(), 'YYYYMMDD') AND
            (SUBSTRING(task_id, 13, 17))='_atf'
    """
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def generateTaskIDDisposal(destination=False):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT
        CASE
            WHEN (MAX(SUBSTRING(task_id, 0, 13))::BIGINT || '_dis') IS NULL THEN CONCAT(TO_CHAR(NOW(), 'YYYYMMDD'), '0001_dis')
            WHEN (MAX(SUBSTRING(task_id, 0, 13))::BIGINT || '_dis') IS NOT NULL THEN CONCAT(MAX(SUBSTRING(task_id, 0, 13))::BIGINT+1, '_dis')
        END as generated_task_id
        FROM tbl_task
        WHERE
            (SUBSTRING(task_id, 0, 9))=TO_CHAR(NOW(), 'YYYYMMDD') AND
            (SUBSTRING(task_id, 13, 17))='_dis'
    """
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getLongLat(site_id):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT lang, lat
        FROM tbl_site
        WHERE site_id = '%s'
    """ % site_id
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getNEID(site_id, ne_code):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT
        CASE
            WHEN COUNT(*) = 0 THEN '{sid}BE{nc}01'
            WHEN RIGHT(MAX(ne_id), 2)::INT<9 THEN LEFT(MAX(ne_id), 13) || '0' || RIGHT(MAX(ne_id), 2)::INT+1
            WHEN RIGHT(MAX(ne_id), 2)::INT>=9 THEN LEFT(MAX(ne_id), 13) || RIGHT(MAX(ne_id), 2)::INT+1
        END as generated_ne_id
        FROM tbl_task
        WHERE LEFT(ne_id , 13) = '{sid}BE{nc}'
    """.format(sid=site_id, nc=ne_code)
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getDivision(vendor):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT item
        FROM tbl_division
        WHERE item ~* '%s'
    """ % (vendor + ' dismantle' if 'mpfs' not in vendor else '')
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def insertTblTask(
        task_id,
        project_id,
        site_id,
        plan_date,
        status,
        remark,
        task_type,
        last_update,
        vendor,
        ne_id,
        atf_no,
        division,
        site_id_destination,
        creator_id=" "
):
    # ===========================================
    # UPLOAD TO tbl_task
    # ===========================================

    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    sql = """
        INSERT INTO tbl_task(
            task_id,
            project_id,
            site_id,
            plan_date,
            status,
            remark,
            task_type,
            last_update,
            vendor,
            ne_id,
            atf_no,
            division,
            site_id_destination,
            creator_id
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """

    cursor.execute(sql, (
        task_id, project_id, site_id, plan_date, status, remark,
        task_type, last_update, vendor, ne_id, atf_no, division,
        site_id_destination, creator_id
    ))
    connection.commit()

    cursor.close()
    connection.close()


def insertTblTaskDestination(
        task_id,
        project_id,
        site_id,
        plan_date,
        status,
        remark,
        task_type,
        last_update,
        vendor,
        ne_id,
        atf_no,
        division,
        site_id_destination,
        site_id_ori,
        task_id_reference
):
    # ===========================================
    # UPLOAD TO tbl_task
    # ===========================================

    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    sql = """
        INSERT INTO tbl_task(
            task_id,
            project_id,
            site_id,
            plan_date,
            status,
            remark,
            task_type,
            last_update,
            vendor,
            ne_id,
            atf_no,
            division,
            site_id_destination,
            site_id_ori,
            task_id_reference
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """

    cursor.execute(sql, (
        task_id, project_id, site_id, plan_date, status, remark,
        task_type, last_update, vendor, ne_id, atf_no, division,
        site_id_destination, site_id_ori, task_id_reference
    ))
    connection.commit()

    cursor.close()
    connection.close()


def insertTblItem(
        pnx_major_eqp,
        pnx_manufacturing,
        pnx_eqp_category,
        pnx_equipment_type,
        pnx_unit_category,
        pnx_serial_num,
        unit_of_measurement,
        qty_waste,
        flag_waste,
        last_update,
        moving_status,
        pnx_operation_status,
        task_id,
        item_id
):
    # ===========================================
    # UPLOAD TO tbl_item
    # ===========================================

    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    sql = """
        INSERT INTO tbl_item(
            pnx_major_eqp,
            pnx_manufacturing,
            pnx_eqp_category,
            pnx_equipment_type,
            pnx_unit_category,
            pnx_serial_num,
            unit_of_measurement,
            qty_waste,
            flag_waste,
            last_update,
            moving_status,
            pnx_operation_status,
            task_id,
            item_id
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """

    cursor.execute(sql, (
        pnx_major_eqp, pnx_manufacturing, pnx_eqp_category, pnx_equipment_type, pnx_unit_category,
        pnx_serial_num, unit_of_measurement, qty_waste, flag_waste, last_update, moving_status,
        pnx_operation_status, task_id, item_id
    ))
    connection.commit()

    cursor.close()
    connection.close()


def getATFNo(atfno):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT COUNT(*)
        FROM tbl_task
        WHERE atf_no = '%s'
    """ % atfno
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def generateItemID(dt):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT
        CASE
            WHEN (MAX(SUBSTRING(item_id, 0, 19))::BIGINT) IS NULL THEN CONCAT(TO_CHAR(NOW(), 'YYYYMMDDHH24MISS'), '0001')
            WHEN (MAX(SUBSTRING(item_id, 0, 19))::BIGINT) IS NOT NULL THEN CONCAT(MAX(SUBSTRING(item_id, 0, 19))::BIGINT+1)
        END as generated_item_id
        FROM tbl_item
        WHERE (SUBSTRING(item_id, 0, 15))='%s'
    """ % dt
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getDummySN():
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT
        CASE
            WHEN COUNT(*) = 0 THEN 'DUMMY-SN-1'
            WHEN COUNT(*) > 0 THEN CONCAT('DUMMY-SN-', (REGEXP_MATCHES(MAX(pnx_serial_num), '^DUMMY-SN-(\d+)$'))[1]::INT+1)
        END as generated_sn
        FROM (
            SELECT pnx_serial_num
            FROM tbl_item
            WHERE pnx_serial_num ~* '^DUMMY\-SN\-\d{1,}'
            ORDER BY SUBSTRING(pnx_serial_num FROM '([0-9]+)')::BIGINT DESC, pnx_serial_num
            LIMIT 1
        ) a
    """
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getPN(system, subsystem, pn):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT kd_acs, category, name, description, vend, type, loc, flag_waste, unit_of_measurement
        FROM tbl_acs
        WHERE category = '%s' AND name ~* '^%s$' AND loc ~* '^%s$'
        ORDER BY id ASC LIMIT 1
    """ % (system, subsystem, pn)
    cursor.execute(query)
    data = cursor.fetchall()

    if not data:
        connection = connect_to_database_netgear()
        cursor = connection.cursor()
        query = """
            SELECT kd_acs, category, name, description, vend, type, loc, flag_waste, unit_of_measurement
            FROM tbl_acs
            WHERE category = '%s' AND loc ~* '^%s$'
            ORDER BY id ASC LIMIT 1
        """ % (system, pn)
        # print(query)
        cursor.execute(query)
        data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getOriginTaskData(task_id_destination):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT task_id, site_id, ne_id, division, site_id_destination, creator_id, vendor
        FROM tbl_task
        WHERE task_id IN (
            SELECT task_id_reference
            FROM tbl_task
            WHERE task_id = '%s'
        )
        ORDER BY id ASC LIMIT 1
    """ % task_id_destination
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getSite(site_id):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT site_id, area
        FROM tbl_site
        WHERE site_id = '%s'
        ORDER BY id ASC LIMIT 1
    """ % site_id
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def getTaskData(task_id):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT task_id, site_id, ne_id, division, site_id_destination, creator_id
        FROM tbl_task
        WHERE task_id = '%s'
        ORDER BY id ASC LIMIT 1
    """ % task_id
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    return data


def insertTblItemFRI(
        pnx_major_eqp,
        pnx_manufacturing,
        pnx_eqp_category,
        pnx_equipment_type,
        pnx_unit_category,
        pnx_serial_num,
        unit_of_measurement,
        qty_waste,
        flag_waste,
        last_update,
        moving_status,
        pnx_operation_status,
        task_id,
        item_id
):
    # ===========================================
    # UPLOAD TO tbl_item FRI
    # ===========================================

    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    sql = """
        INSERT INTO tbl_item_fri(
            pnx_major_eqp,
            pnx_manufacturing,
            pnx_eqp_category,
            pnx_equipment_type,
            pnx_unit_category,
            pnx_serial_num,
            unit_of_measurement,
            qty_waste,
            flag_waste,
            last_update,
            moving_status,
            pnx_operation_status,
            task_id,
            item_id
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """

    cursor.execute(sql, (
        pnx_major_eqp, pnx_manufacturing, pnx_eqp_category, pnx_equipment_type, pnx_unit_category,
        pnx_serial_num, unit_of_measurement, qty_waste, flag_waste, last_update, moving_status,
        pnx_operation_status, task_id, item_id
    ))
    connection.commit()

    cursor.close()
    connection.close()


def getFRIData(task_id):
    connection = connect_to_database_netgear()
    cursor = connection.cursor()
    query = """
        SELECT status, site_id
        FROM tbl_task
        WHERE task_id = '%s'
    """ % (task_id)
    cursor.execute(query)
    data = cursor.fetchall()

    cursor.close()
    connection.close()

    data2 = []

    if 'Approved' in data[0][0]:
        connection = connect_to_database_netgear()
        cursor = connection.cursor()
        query2 = """
            SELECT
                pnx_major_eqp,
                pnx_manufacturing,
                pnx_eqp_category,
                pnx_equipment_type,
                pnx_unit_category,
                pnx_serial_num,
                pnx_mgr_status,
                description,
                pnx_operation_status,
                pnx_mgr_remark,
                '%s' as site_id,
                barcode,
                moving_status,
                unit_of_measurement,
                flag_waste,
                qty_waste,
                qty_waste,
                reuse_matrix,
                pallet_id,
                locator_id,
                condition,
                remark_condition
            FROM tbl_item
            WHERE task_id = '%s'
        """ % (data[0][1], task_id)
        cursor.execute(query2)
        data2 = cursor.fetchall()

        cursor.close()
        connection.close()
    else:
        connection = connect_to_database_netgear()
        cursor = connection.cursor()
        query2 = """
            SELECT
                pnx_major_eqp,
                pnx_manufacturing,
                pnx_eqp_category,
                pnx_equipment_type,
                pnx_unit_category,
                pnx_serial_num,
                pnx_mgr_status,
                description,
                pnx_operation_status,
                pnx_mgr_remark,
                '%s' as site_id,
                barcode,
                moving_status,
                unit_of_measurement,
                flag_waste,
                qty_waste,
                qty_waste,
                reuse_matrix,
                pallet_id,
                locator_id,
                condition,
                remark_condition
            FROM tbl_item_fri
            WHERE task_id = '%s'
        """ % (data[0][1], task_id)
        cursor.execute(query2)
        data2 = cursor.fetchall()

        cursor.close()
        connection.close()

    return data2

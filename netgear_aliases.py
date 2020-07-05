def getSystemAlias(system):
    return {
        'BSS (RAN)': 'LTE',
        'CME': 'CME',
        'CORE': 'HLR',
        'TRM': 'TPM',
        'WLAN': 'LTE',
        'TPM': 'TPM'
    }[system]


def getWHIDAlias(whid):
    return {
        'WH01': '03WDHL02',  # WH CIBARUSAH
        'WH02': '20WDHL01',  # WH DHL Surabaya Tambak
        'WH03': '11WDHL01',  # WH DHL Palembang Alang Blok A 04
        'WH04': '07WDHL01',  # WH DHL Batam Yossudarso
        'WH05': '28WDHL01',  # WH DHL Makassar IR Sutami
        'WH06': '19WDHL01',  # WH_DHL_BALIKPAPAN_PROJAKAL
        'WH07': '05WDHL01',  # WH DHL Medan Komplek ATC
        'WH08': '16WDHL01',  # WH_DHL_PONTIANAK_ADISUCIPTO
        'WH09': '14WDHL01',  # WH DHL Semarang GS Blok 23
        'WH 01': '03WDHL02',  # WH CIBARUSAH
        'WH 02': '20WDHL01',  # WH DHL Surabaya Tambak
        'WH 03': '11WDHL01',  # WH DHL Palembang Alang Blok A 04
        'WH 04': '07WDHL01',  # WH DHL Batam Yossudarso
        'WH 05': '28WDHL01',  # WH DHL Makassar IR Sutami
        'WH 06': '19WDHL01',  # WH_DHL_BALIKPAPAN_PROJAKAL
        'WH 07': '05WDHL01',  # WH DHL Medan Komplek ATC
        'WH 08': '16WDHL01',  # WH_DHL_PONTIANAK_ADISUCIPTO
        'WH 09': '14WDHL01',  # WH DHL Semarang GS Blok 23
        'DHL01': '03WDHL01',
        'DHL02': '20WDHL01',
        'DHL03': '11WDHL01',
        'DHL04': '07WDHL01',
        'DHL05': '28WDHL01',
        'DHL06': '19WDHL01',
        'DHL07': '05WDHL01',
        'DHL08': '16WDHL01',
        'DHL09': '14WDHL01',
        'WH MML': '05WMML01',
        'WAREHOUSE NEC PONTIANAK': '16WNEC01'
    }[whid]


def getVendorAlias(vendor):
    return {
        'ZTE INDONESIA': 'PT.ZTE Indonesia',
        'ZTE INDONESA': 'PT.ZTE Indonesia',
        'NEC': 'NEC Indonesia'
    }[vendor]


def getDivisionAlias(area):
    divdict = {
        'JABODETABEK': 'RAN Jabodetabek',
        'WEST JAVA': 'RAN Java & Bali Nusa',
        'CENTRAL JAVA AND DIY': 'RAN Java & Bali Nusa',
        'EAST JAVA AND BALI': 'RAN Java & Bali Nusa',
    }
    return divdict[area] if area in divdict else 'RAN & Transport Access Program Mgt.'

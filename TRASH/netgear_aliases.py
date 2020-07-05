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
        'WH01': '03WDHL01',
        'WH02': '20WDHL01',
        'WH03': '11WDHL01',
        'WH04': '07WDHL01',
        'WH05': '28WDHL01',
        'WH06': '19WDHL01',
        'WH07': '05WDHL01',
        'WH08': '16WDHL01',
        'WH09': '14WDHL01',
        'DHL01': '03WDHL01',
        'DHL02': '20WDHL01',
        'DHL03': '11WDHL01',
        'DHL04': '07WDHL01',
        'DHL05': '28WDHL01',
        'DHL06': '19WDHL01',
        'DHL07': '05WDHL01',
        'DHL08': '16WDHL01',
        'DHL09': '14WDHL01'
    }[whid]

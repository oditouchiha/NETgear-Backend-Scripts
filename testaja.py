
import glob

list_of_atf_noarr = glob.glob("%s\\*.xlsx" % "D:\\KERJAAN\\SCRIPTS\\SCRIPTS\\Made by AMG\\MANUAL - NETgear ATF Uploader\\INPUT\\NOARR")

for atf_file in list_of_atf_noarr:
    if '~$' not in atf_file:
        print(atf_file)

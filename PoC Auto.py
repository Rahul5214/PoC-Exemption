import csv
import datetime
import colorama     # pip3 install colorama

from colorama import Fore,Style,Back
colorama.init()

currentDate = datetime.datetime.now()
currentYear = currentDate.strftime("%y")        # Returns "22" for year 2022

print(f"{Fore.YELLOW}Enter DATE and MONTH only (Day Ahead) DDMM: ", end='')
date = input()

PoC_Status = {
    "Injecting Utility" : "PoC Exemption",
    "AHEJ2L_SOLAR"   :"N",
    "AHEJ2L_S_FTG2"  :"Y",
    "AHEJ2L_W_FTG2"	 :"Y",
    "AHEJ3L_SOLAR"	 :"N",
    "AHEJ3L_S_FTG2"	 :"Y",
    "AHEJ3L_W_FTG2"	 :"Y",
    "AHEJ4L_SOLAR"	 :"N",
    "AHEJ4L_S_FTG1"	 :"N",
    "AHEJ4L_W_FTG1"	 :"N",
    "AHEJOL_SOLAR"	 :"N",
    "AHEJOL_S_FTG2"	 :"Y",
    "AHEJOL_W_FTG2"	 :"Y",
    "ARP1PL_BKN"	 :"N",
    "ASEJ2L"	     :"N",
    "ARERJL"	     :"Y",
    "AP43PL_BKN"	 :"N",
    "RSEJ3PL_FTG2"	 :"N",
    # "RSUPL_FTG2"	 :"Y",   # "Y" for Telangana, "N" for others
    "TPREL"	         :"Y",
    "RSRPL_BKN"	     :"N",
    "APMPL_BHDL"	 :"N",
    "AvSusRJPPL_BKN" :"N",
    "TS1PL_BKN"	     :"Y",
    "CSPJPL_BHDL"	 :"Y",
    "AcHPPL_BHDL2"	 :"Y",
    "ABCREPL_BHDL2"	 :"Y",
    "MSUPL_BHDL2"	 :"Y",
    "TPGEL_BKN"      :"N",
    "KSPPNL_BHDL2"   :"Y",
}

POC = []

def automate_POC():
    try:
        with open("IMPSCH" + date + currentYear + ".csv", "r", newline="") as file:
            reader = csv.reader(file)

            allRows = list(reader)         # Entire file is stored in this variable in 2D array form
            fifthrow = allRows[4]          # Contains all the utilities
            sixthrow = allRows[5]          # For checking if state is Telangana or not for "RSUPL_FTG2"
            
            if(allRows[12][0] == "PoC Exemption"):
                print(f"{Fore.GREEN}PoC status already exist in the Excel file")
                print(f"{Fore.GREEN}Success")
                input ("Press Enter to close")
                return

            for index, utility in enumerate(fifthrow):              # enumerate is used to get index of utility "RSUPL_FTG2"
                try:
                    if utility == "RSUPL_FTG2":
                        if sixthrow[index] == "Telangana":
                            POC.append("Y")
                        else:
                            POC.append("N")
                    else:
                        POC.append(PoC_Status[utility])
                
                except KeyError as e:
                    print (f"\n{Fore.RED}Error: PoC Exemption status is not defined for utility, {e}")
                    POC.append("NOT DEFINED")
            
            allRows.insert(12,POC)      # for inserting "Y" and "N" in 13th row in excel

        try:
            with open("IMPSCH" + date + currentYear + ".csv", "w", newline="") as file:
                writer = csv.writer(file)
                writer.writerows(allRows)
        except PermissionError as e:
            print (f"{Fore.RED}Error: Excel file is open. Please close it before executing python script, {e}")
            input ("Press Enter to close")
            return

    except FileNotFoundError as e:
        print (f"{Fore.RED}{e}")
        input ("Press Enter to close")

    else:
        print (f"{Fore.GREEN}Success")
        input ("Press Enter to close")


automate_POC()

# This is a sample Python script.

# Press Mayús+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import os
import subprocess
import xlsxwriter
import sys


def main():
    bit_4_range = [2, 255]

    actives = 0
    inactives = 0

    excel = xlsxwriter.Workbook('excel.xlsx')
    worksheet = excel.add_worksheet()

    index = 1

    with open(os.devnull, "wb") as limbo:
        try:
            for bit_4 in range(bit_4_range[0], bit_4_range[1] + 1):

                ip = '192.168.1.{}'.format(bit_4)
                result = subprocess.Popen(["ping", "-n", "1", "-w", "300", ip],
                                          stdout=limbo, stderr=limbo).wait()
                if result:
                    status = "inactive"
                    inactives = inactives + 1
                else:
                    status = "active"
                    actives = actives + 1

                print(ip, status)
                worksheet.write('A{}'.format(index), ip)
                worksheet.write('B{}'.format(index), status)
                index = index + 1

        except KeyboardInterrupt:
            print('Cerrando forzosamente...')

        finally:
            print("TOTAL..............")
            print("Activos", actives)
            print("Inactivos", inactives)
            excel.close()
            sys.exit()


if __name__ == '__main__':
    main()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/

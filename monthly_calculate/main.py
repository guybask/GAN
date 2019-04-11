import xlrd
import csv
import datetime
from calendar import monthrange
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

xlsFile = "/Users/guybaskin/myProj/GAN/output/apts_summary.xlsx"
csvTargetFile = "/Users/guybaskin/myProj/GAN/output/apt_summary"
outputFile = "/Users/guybaskin/myProj/GAN/output/output_data.xlsx"

CURRENT_YEAR    = 2018

class Financial :
    def __init__(self,month,year,apartment):
        self.month  = month
        self.year   = year
        self.income = Income(0,0,0)
        self.booking_commition  = 0
        self.fixed_expenses = apartment.arnona + apartment.cable_tv + apartment.electricity + apartment.rent
        self.cleaning = 0
        self.net    = 0
        self.others = 0
        self.apartment = apartment

class TableHeaders :
    NUMBER_OF_INPUT_PARMAS = 12
    NAME, CHECK_IN, CHECK_OUT, TOTAL_OF_DAYS, CASH, CREDIT, TOTAL_PAYMENT, CLEANING_FEE, OTHERS, TOTAL_PAYMENT2, IS_BOOKING, EARNING_FROM_COMMITION = range (NUMBER_OF_INPUT_PARMAS)

class OutputParams :
    NUMBER_OF_OUTPUT_PARAMS = 8
    TOTAL_INCOME, CREDIT_INCOME, CASH_INCOME, FIXED_EXPENSES, CLEANING, BOOKING_COMMTION, OTHERS, NET = range(NUMBER_OF_OUTPUT_PARAMS)

    _display_strings = {
        TOTAL_INCOME : "Income",
        CREDIT_INCOME : "Credit Income",
        CASH_INCOME: "Cash Income",
        FIXED_EXPENSES: "Fixed Expenses ( rent + arnona + cables + electricity ) ",
        CLEANING: "Cleaning",
        BOOKING_COMMTION: "Booking Commition",
        OTHERS: "Others",
        NET: "NET"
    }

class Excel_manager :
    def __init__(self,file_name):
        self.workbook = xlsxwriter.Workbook(file_name)
        self.worksheet = self.workbook.add_worksheet("Summary")
        self.red_cell_format    = self.workbook.add_format()
        self.green_cell_format  = self.workbook.add_format()
        self.yellow_cell_format = self.workbook.add_format()
        self.gold_cell_format   = self.workbook.add_format()
        self.red_cell_format.set_bg_color('red')
        self.yellow_cell_format.set_bg_color('yellow')
        self.green_cell_format.set_bg_color('green')
        self.gold_cell_format.set_bg_color('orange')
        self.cell_fomat_list = [self.yellow_cell_format, self.red_cell_format, self.green_cell_format, self.gold_cell_format]

        self.year_format = self.workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'grey'})
        cell = xlsxwriter

class amount :
    def __init__(self,online_platforms,friends,total):
        self.online_platforms = online_platforms
        self.friends = friends
        self.total = total


class Reservation :
    def __init__(self):
        self.credit         = 0
        self.cash           = 0
        self.total_payment  = 0
        self.cleaning       = 0
        self.num_of_days    = 0
        self.credit_or_cash = ''
        self.check_in       = 0
        self.check_out      = 0
        self.booking_commition = 0
        self.earning_from_commition = 0

    def build_reservation_from_raw(self,raw):
        #print raw
        raw_tablbe = raw.split(',')
        for i,item in enumerate(raw_tablbe) :
            raw_tablbe[i] = item.strip('"')
            raw_tablbe[i] = raw_tablbe[i].strip('\r\n')
            raw_tablbe[i] = raw_tablbe[i].strip('"')
        if raw_tablbe[TableHeaders.NAME] == '':
            return (1)
        #try :
        self.total_payment  = float(raw_tablbe[TableHeaders.TOTAL_PAYMENT])
        self.cleaning       = float(raw_tablbe[TableHeaders.CLEANING_FEE])
        self.num_of_days    = float(raw_tablbe[TableHeaders.TOTAL_OF_DAYS])
        self.credit_or_cash = ''
        self.check_in       = float_to_date(float(raw_tablbe[TableHeaders.CHECK_IN]))
        self.check_out      = float_to_date(float(raw_tablbe[TableHeaders.CHECK_OUT]))
        if raw_tablbe[TableHeaders.CREDIT] == '':
            self.credit = 0
        else:
            self.credit = float(raw_tablbe[TableHeaders.CREDIT])
        if raw_tablbe[TableHeaders.CASH] == '':
            self.cash = 0
        else :
            self.cash = float(raw_tablbe[TableHeaders.CASH])
        if (raw_tablbe[TableHeaders.IS_BOOKING] == 'x' ):
            print raw_tablbe
            self.booking_commition = float(raw_tablbe[TableHeaders.OTHERS])
        else :
            self.booking_commition = 0
        if (raw_tablbe[TableHeaders.EARNING_FROM_COMMITION] != '' )and (raw_tablbe[TableHeaders.EARNING_FROM_COMMITION] != '\r\n' ):
            print raw_tablbe
            self.earning_from_commition = float(raw_tablbe[TableHeaders.EARNING_FROM_COMMITION])
        else :
            self.earning_from_commition = 0
        # except Exception:
        #     print "recieved exception "

class Expenses :
    NUMBER_OF_TABLE_PARAMS = 4
    ITEM, DATE, PRICE, APARTMENT = range(NUMBER_OF_TABLE_PARAMS)
    def __init__(self):
        self.date = ''
        self.amount = 0
        self.apartment = ''
        self.item = ''

    def build_expenses_from_raw(self,raw):
        raw_tablbe = raw.split(',')
        for i, item in enumerate(raw_tablbe):
            raw_tablbe[i] = item.strip('"')
            raw_tablbe[i] = raw_tablbe[i].strip('\r\n')
            raw_tablbe[i] = raw_tablbe[i].strip('"')
        if raw_tablbe[Expenses.ITEM] == '':
            return (1)
        print "debubg",raw_tablbe[Expenses.PRICE]
        self.amount         = float(raw_tablbe[Expenses.PRICE])
        self.date = float_to_date(float(raw_tablbe[Expenses.DATE]))
        self.apartment = raw_tablbe[Expenses.APARTMENT]

class rent :
    def __init__(self,contract, cash, total):
        self.contract = contract
        self.cash = cash
        self.total = total

class Income :
    def __init__(self,credit , cash, total):
        self.credit = credit
        self.cash = cash
        self.total = total

    def add_to_sum(self,credit,cash,total):
        self.credit += credit
        self.cash   += cash
        self.total  += total

class Apartment :
    def __init__(self,name,rent,electricity,cable_tv,arnona):
        self.name           = name
        self.rent           = rent
        self.electricity    = electricity
        self.cable_tv       = cable_tv
        self.arnona         = arnona


def float_to_date(float_date):
    seconds = (float_date - 25569) * 86400.0
    return datetime.datetime.utcfromtimestamp(seconds)

class GordonManager :
    def __init__(self):
        self.apartment_list = []
    # Currently hard codded -> need to read this from excel file
    #============================================================================================================================#
        apartment_yellow = Apartment('yellow',6500,100,110,100)
        apartment_red    = Apartment('red', 6500, 100, 110, 100)
        apartment_green  = Apartment('green', 7000, 180, 150, 200)
        apartment_gold = Apartment('gold', 0,0,0,0)
    # ============================================================================================================================#
        self.apartment_list.append(apartment_yellow)
        self.apartment_list.append(apartment_red)
        self.apartment_list.append(apartment_green)
        self.apartment_list.append(apartment_gold)

    ######################################################
    ########## INPUT    : Path of Excel file
    ########## OUTPUT   : list of files, each file contains list of csv rows
    ######################################################
    def create_csv_from_xlsx(self, path):
        csv_list = []
        wb = xlrd.open_workbook(path)
        for apartment in self.apartment_list :
            sh = wb.sheet_by_name(apartment.name)
            your_csv_file = open(csvTargetFile + '_' + apartment.name + '.csv', 'wb')

            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
            for rownum in xrange(sh.nrows):
                try:
                    wr.writerow(sh.row_values(rownum))
                except Exception:
                    print rownum
                    print apartment.name
            csv_list.append(csv.reader(your_csv_file))
            your_csv_file.close()


        sh = wb.sheet_by_name('expenses')
        your_csv_file = open(csvTargetFile + '_' + 'expenses' + '.csv', 'wb')

        wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
        for rownum in xrange(sh.nrows):
            try:
                wr.writerow(sh.row_values(rownum))
            except Exception:
                print rownum
                print apartment.name
        csv_list.append(csv.reader(your_csv_file))
        your_csv_file.close()


    ######################################################
    ########## INPUT    : list of years and month to return report for
    ########## OUTPUT   : list of year, each contains list of month, each contains list Financial
    ######################################################
    def get_income_by_month(self,month_list,year_list = CURRENT_YEAR):

        year_output_list = [] # list of month
        for year in year_list :
            month_output_list = [] # list of apartments
            for month in month_list:
                month_summary = []
                for i,apt in enumerate(self.apartment_list):
                    apt_monthly_summary = Financial(month,year,apt)
                    apt_monthly_summary.income, apt_monthly_summary.cleaning, apt_monthly_summary.booking_commition = self.calculate_income_and_cleaning_from_csv_per_month(apt.name,year,month)
                    apt_monthly_summary.others = self.calculate_expenses_from_csv('expenses',year,month,apt.name) + ( self.calculate_expenses_from_csv('expenses',year,month,'ALL') / len(self.apartment_list))
                    apt_monthly_summary.net = apt_monthly_summary.income.total - apt_monthly_summary.cleaning - apt_monthly_summary.booking_commition - apt_monthly_summary.fixed_expenses - apt_monthly_summary.others
                    month_summary.append(apt_monthly_summary)
                month_output_list.append(month_summary)
            year_output_list.append(month_output_list)
        return year_output_list

    # ======================================================================================================================================================#
    def calculate_expenses_from_csv(self,csv_file_name,year,month,apt_name):
        csv_file = open(csvTargetFile + '_' + csv_file_name+ '.csv', 'r')
        sum = 0
        for i,raw in enumerate(csv_file) :
            partial_sum = 0
            if i == 0   :
                continue
            new_expense = Expenses()
            return_code = new_expense.build_expenses_from_raw(raw)
            if return_code == 1 :
                break
            if (new_expense.date.year == year) and (new_expense.date.month == month) :
                if (new_expense.apartment.lower() == apt_name.lower())  :
                    partial_sum = new_expense.amount
            sum = sum + partial_sum
        return sum
    #======================================================================================================================================================#
    def calculate_income_and_cleaning_from_csv_per_month(self,csv_file_name,year,month):
        csv_file = open(csvTargetFile + '_' + csv_file_name + '.csv', 'r')
        sum = Income(0,0,0)
        cleaning = 0
        booking_commition_sum = 0
        Reservation_list_for_month = []
        for i,raw in enumerate(csv_file) :
            partial_sum = Income(0,0,0)
            partial_sum_cleaning = 0
            reservation_booking_commition = 0
            # skip the first line of headers
            if i == 0   :
                continue
            newReservation = Reservation()
            returnCode = newReservation.build_reservation_from_raw(raw)
            if returnCode == 1 :
                break
            monthRange = monthrange(newReservation.check_in.year,month)
            # TODO : improve algorithm to calculate revenues for the same month in case it is spread between more than one month
            # reservation begins and ends on the same year :
            if (newReservation.check_in.year == year)  and (newReservation.check_out.year == year) :
                # reservation begin and ends on the same nonth :
                if (newReservation.check_in.month == month) and (newReservation.check_out.month == month ) :
                    partial_sum.total               = newReservation.total_payment
                    partial_sum.cash                = newReservation.cash
                    partial_sum.credit              = newReservation.credit
                    partial_sum_cleaning            = newReservation.cleaning
                    reservation_booking_commition   = newReservation.booking_commition
                elif (newReservation.check_in.month == month) or (newReservation.check_out.month == month ) :
                    # reservation begins on this month, but ends on another one :
                    if (newReservation.check_in.month == month) :
                        partial_sum.total               = ((newReservation.total_payment)/ newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum.cash                = ((newReservation.cash) / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum.credit              = ((newReservation.credit) / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum_cleaning            = (newReservation.cleaning / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        reservation_booking_commition   = (newReservation.booking_commition / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        # reservation ends on this month :
                    elif (newReservation.check_out.month == month):
                        partial_sum.total               = ((newReservation.total_payment)/ newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum.cash = ((newReservation.cash) / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum.credit = ((newReservation.credit) / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum_cleaning            = (newReservation.cleaning / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        reservation_booking_commition   = (newReservation.booking_commition / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                # reservation started before month and ended after
                elif (newReservation.check_in.month < month) and (newReservation.check_out.month > month ) :
                    partial_sum.total                     = ((newReservation.total_payment)/ newReservation.num_of_days) * monthRange[1]
                    partial_sum.cash                      = ((newReservation.cash) / newReservation.num_of_days) * monthRange[1]
                    partial_sum.credit                     = ((newReservation.credit) / newReservation.num_of_days) * monthRange[1]
                    partial_sum_cleaning            = (newReservation.cleaning / newReservation.num_of_days) * monthRange[1]
                    reservation_booking_commition   = (newReservation.booking_commition / newReservation.num_of_days) * monthRange[1]
                else :
                    continue
            # reservation begin on one year and ends on next one
            elif (newReservation.check_in.year == year)  or (newReservation.check_out.year == year) :
                if (newReservation.check_in.month == month) or (newReservation.check_out.month == month):
                    # reservation begins on this month, but ends on another one :
                    if (newReservation.check_in.month == month) and (newReservation.check_in.year == year) :
                        partial_sum.total                     = ((newReservation.total_payment)/ newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum.cash = ((newReservation.cash) / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum.credit = ((newReservation.credit) / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        partial_sum_cleaning            = (newReservation.cleaning / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        reservation_booking_commition   = (newReservation.booking_commition / newReservation.num_of_days) * (monthRange[1] - newReservation.check_in.day + 1)
                        # reservation ends on this month :
                    elif (newReservation.check_out.month == month) and (newReservation.check_out.year == year):
                        partial_sum.total                    = ((newReservation.total_payment) / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum.cash                     = ((newReservation.cash) / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum.credit                   = ((newReservation.credit) / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        partial_sum_cleaning            = ( newReservation.cleaning  / newReservation.num_of_days) * (newReservation.check_out.day - 1)
                        reservation_booking_commition   = ( newReservation.booking_commition  / newReservation.num_of_days) * (newReservation.check_out.day - 1)
            else :
                continue
            # add partial sum to sum
            if newReservation.earning_from_commition != 0 :
                ratio = partial_sum.total / newReservation.total_payment
                partial_sum = newReservation.earning_from_commition * ratio + partial_sum_cleaning
                sum.add_to_sum(0,partial_sum,partial_sum)
            else :
                sum.add_to_sum(partial_sum.credit,partial_sum.cash,partial_sum.total)
            cleaning+= partial_sum_cleaning
            booking_commition_sum+= reservation_booking_commition
            Reservation_list_for_month.append(newReservation)
        print "month %d, SUM %d "%(month,sum.total)
        return sum, cleaning, booking_commition_sum

    # ======================================================================================================================================================#

    def create_table_structur(self,output_data,excel_manager) :
        # ===================
        # CREATE YEAR STRUCTURE
        number_of_month = output_data[0].__len__()
        col = 0
        row = 2
        for i,year in enumerate(output_data):
            excel_manager.worksheet.write((row+(i * number_of_month)),col,year[0][0].year)
            for j,month in enumerate(year) :
                excel_manager.worksheet.write((row + (i * number_of_month) + j), col+1, month[0].month)
            row = row + 1
        # ===================
        # CREATE APT STRUCTURE
        # number_of_apartments = output_data[0][0].__len__()
        col = 2
        row = 0
        for i,financial in enumerate(output_data[0][0]) :
            excel_manager.worksheet.write(row,col+(i * int(OutputParams.NUMBER_OF_OUTPUT_PARAMS)),financial.apartment.name,excel_manager.cell_fomat_list[i])
            for j in range(OutputParams.NUMBER_OF_OUTPUT_PARAMS) :
                excel_manager.worksheet.write(row + 1,col + (i * OutputParams.NUMBER_OF_OUTPUT_PARAMS) + j, OutputParams._display_strings[j])


    # ======================================================================================================================================================#

    def export_data_to_xlsx(self,output_data):
        # Create a workbook and add a worksheet.
        excel_manager = Excel_manager(outputFile)
        number_of_month = output_data[0].__len__()

        self.create_table_structur(output_data,excel_manager)
        row = 2
        for i,year in enumerate(output_data):
            year_row = row + i * number_of_month
            year_col = 2
            row = row + 1
            for j,month in enumerate(year):
                month_row = year_row + j
                month_col = year_col
                for k,financial_report in enumerate(month):
                    data_row = month_row
                    data_col = month_col + k * OutputParams.NUMBER_OF_OUTPUT_PARAMS
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.TOTAL_INCOME, (int)(financial_report.income.total), excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.CREDIT_INCOME,   (int)(financial_report.income.credit),excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.CASH_INCOME, (int)(financial_report.income.cash),excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.FIXED_EXPENSES, (int)(financial_report.fixed_expenses), excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.CLEANING, (int)(financial_report.cleaning), excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.BOOKING_COMMTION, (int)(financial_report.booking_commition), excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.OTHERS, (int)(financial_report.others), excel_manager.cell_fomat_list[k])
                    excel_manager.worksheet.write(data_row, data_col + OutputParams.NET, (int)(financial_report.net) ,excel_manager.cell_fomat_list[k])

        #==================================================
                xl_rowcol_to_cell
                excel_manager.worksheet.write(data_row,data_col + OutputParams.NUMBER_OF_OUTPUT_PARAMS,'=SUM(%s,%s,%s,%s)' %(xl_rowcol_to_cell(data_row,data_col),xl_rowcol_to_cell(data_row,data_col-(OutputParams.NUMBER_OF_OUTPUT_PARAMS)),xl_rowcol_to_cell(data_row,data_col-(2*OutputParams.NUMBER_OF_OUTPUT_PARAMS)),xl_rowcol_to_cell(data_row,data_col-(3*OutputParams.NUMBER_OF_OUTPUT_PARAMS))))

        excel_manager.workbook.close()

if __name__ == '__main__':
    # csv_from_excel('yellow')
    # csv_from_excel('red')
    # csv_from_excel('green')
    manager = GordonManager()
    manager.create_csv_from_xlsx(xlsFile)
    year = [2017,2018]
    month = [1,2,3,4,5,6,7,8,9,10,11,12]

    financial_list = manager.get_income_by_month(month,year)
    for year in financial_list :
        print "YEAR : %s" %str(year[0][0].year)
        for month in year :
            print "    MONTH : %s" %str(month[0].month)
            total_per_month = 0
            for financial_report in month :
                print "        APT : %s"%financial_report.apartment.name
                print "            INCOME : %d"%financial_report.income.total
                print "            CLEANING : %d"%financial_report.cleaning
                print "            EXPENSES : %d" % financial_report.fixed_expenses
                print "            BOOKING COMMITION : %d"%financial_report.booking_commition
                print "            NET INCOME : %d" %financial_report.net
                total_per_month = financial_report.net + total_per_month

    manager.export_data_to_xlsx(financial_list)

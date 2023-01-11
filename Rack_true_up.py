import os, logging, glob, time, sys
from datetime import datetime, date,timedelta
import xlwings as xw
import pandas as pd
import xlwings.constants as win32c
import bu_alerts


def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        raise e

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e


def insert_all_borders(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
        # linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        # for values in linestylevalues:
        a=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop)
        a.LineStyle = win32c.LineStyle.xlContinuous
        a.ColorIndex = 0
        a.TintAndShade = 0
        a.Weight = win32c.BorderWeight.xlThin
        b=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeBottom)
        b.LineStyle = win32c.LineStyle.xlDouble
        b.ColorIndex = 0
        b.TintAndShade = 0
        b.Weight = win32c.BorderWeight.xlThick
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
    except Exception as e:
        raise e

def conditional_formatting(columnvalue:str,working_sheet,working_workbook):
    try:
        font_colour = -16383844
        Interior_colour = 13551615
        working_sheet.api.Range(f"{columnvalue}:{columnvalue}").Select()
        working_workbook.app.selection.api.FormatConditions.AddUniqueValues()
        working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()

        working_workbook.app.selection.api.FormatConditions(1).DupeUnique = win32c.DupeUnique.xlDuplicate

        working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic
        return font_colour,Interior_colour
    except Exception as e:
        raise e


def rackTrueup(priceInput,rackInput,trueup_file,rackOutput):
    try:
        for file in glob.glob(rackInput+"\\*.xlsx"):
            path, file_name = os.path.split(file)
            #Getting prev month dates
            file_date = file_name.split('_')[-1].replace(".xlsx","").strip()
            file_month = datetime.strptime(file_date,"%m.%Y").strftime("%b")
            file_year = datetime.strptime(file_date,"%m.%Y").strftime("%Y")
            file_month2 = datetime.strptime(file_date,"%m.%Y").strftime("%m")
            logging.info("Opening operating workbook instance of excel")
            if os.path.exists(file):
                wb = xlOpner(file)    
            Open_gr_sheet = wb.sheets[f"Open GR {file_month} {file_year}"]
            Open_gr_sheet.activate()
            column_list = Open_gr_sheet.range("B6").expand('right').value
            Voucher_no_column=column_list.index('Voucher')+2
            Voucher_letter_column = num_to_col_letters(Voucher_no_column)

            Open_gr_sheet.api.Range(f"{Voucher_letter_column}6").AutoFilter(Field:=f"{Voucher_no_column-1}", Criteria1:=["=*PVI*"], Operator:=win32c.AutoFilterOperator.xlAnd)


            last_row = Open_gr_sheet.range(f'B'+ str(Open_gr_sheet.cells.last_cell.row)).end('up').row
            last_column = Open_gr_sheet.range('B6').end('right').last_cell.column
            last_column_letter=num_to_col_letters(last_column)

            Open_gr_sheet.api.Range(f"{Voucher_letter_column}6:{last_column_letter}{last_row}").SpecialCells(12).Select()
            # time.sleep(10000)

            wb.app.selection.copy()

            time.sleep(1)
            wb.sheets.add(f"PVI Data {file_month}",after=Open_gr_sheet)
            time.sleep(1)

            PVI_sheet = wb.sheets[f"PVI Data {file_month}"]
            PVI_sheet.range(f"A1").paste()
            Open_gr_sheet.api.AutoFilterMode=False
            wb.app.api.CutCopyMode=False
            PVI_sheet.autofit()
            path= priceInput + f"\\{file_month}{file_year}Prices.xlsx"
            if not os.path.exists(path):
                print(f"{path} Excel file not present for date {file_month}{file_year}")
            PRICING_DF = pd.read_excel(priceInput + f"\\{file_month}{file_year}Prices.xlsx")
            Pricing_index_dict=dict(zip(PRICING_DF[PRICING_DF.columns[0]], PRICING_DF[PRICING_DF.columns[1]])) 


            path2 = trueup_file + f"\\{file_month2}{file_year} AP PO.xlsx"
            if not os.path.exists(path2):
                print(f"{path2} Excel file not present for date {file_month2}{file_year}")

            TRUE_UP_DF = pd.read_excel(trueup_file + f"\\{file_month2}{file_year} AP PO.xlsx")
            TRUE_UP_index_dict=dict(zip(TRUE_UP_DF[TRUE_UP_DF.columns[0]], TRUE_UP_DF[TRUE_UP_DF.columns[1]])) 
            em_df = pd.DataFrame(columns = ['Vendor', 'Location', 'Qty', 'Amount', 'Diff', 'Pricing Terms'])
            for key,value in TRUE_UP_index_dict.items():
                Open_gr_sheet.activate()
                Links_no_column=column_list.index('Links')+2
                Links_letter_column = num_to_col_letters(Links_no_column)
                Open_gr_sheet.api.Range(f"{Links_letter_column}6").AutoFilter(Field:=f"{Links_no_column-1}", Criteria1:=[f"POR:{value}"], Operator:=win32c.AutoFilterOperator.xlAnd)
                last_row = Open_gr_sheet.range(f'B'+ str(Open_gr_sheet.cells.last_cell.row)).end('up').row
                last_column = Open_gr_sheet.range('B6').end('right').last_cell.column
                last_column_letter=num_to_col_letters(last_column)

                Open_gr_sheet.api.Range(f"{Voucher_letter_column}6:{last_column_letter}{last_row}").SpecialCells(12).Select()
                # time.sleep(10000)

                wb.app.selection.copy()

                time.sleep(1)
                wb.sheets.add(f"{key} MRN {file_month}",after=Open_gr_sheet)
                time.sleep(1)

                CHS_MRN_sheet = wb.sheets[f"{key} MRN {file_month}"]
                CHS_MRN_sheet.range(f"A1").paste()
                Open_gr_sheet.api.AutoFilterMode=False
                wb.app.api.CutCopyMode=False
                CHS_MRN_sheet.autofit()

                last_row_chs = CHS_MRN_sheet.range(f'A'+ str(Open_gr_sheet.cells.last_cell.row)).end('up').row
                PVI_sheet.activate()
                PVI_column_list = PVI_sheet.range("A1").expand('right').value

                Pvi_Links_no_column=PVI_column_list.index('Links')+1
                Pvi_Links_letter_column = num_to_col_letters(Pvi_Links_no_column)
                Pvi_last_row = PVI_sheet.range(f'{Pvi_Links_letter_column}'+ str(PVI_sheet.cells.last_cell.row)).end('up').row

                CHS_MRN_sheet.activate()
                CHS_MRN_sheet.range(f"A2:A{last_row_chs}").copy()
                time.sleep(1)
                PVI_sheet.activate()
                PVI_sheet.range(f'{Pvi_Links_letter_column}{Pvi_last_row+5}').paste()

                font_colour,Interior_colour = conditional_formatting(columnvalue=Pvi_Links_letter_column,working_sheet=PVI_sheet,working_workbook=wb)

                PVI_sheet.api.Range(f"{Pvi_Links_letter_column}1").AutoFilter(Field:=f"{Pvi_Links_no_column}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
                
                Account_no_column=PVI_column_list.index('Account')+1
                Account_letter_column = num_to_col_letters(Account_no_column)
                try:
                    PVI_sheet.api.Range(f"{Account_letter_column}1").AutoFilter(Field:=f"{Account_no_column}", Criteria1:=[key+f"*"])
                except:
                    pass    

                Pvi_last_row = PVI_sheet.range(f'A'+ str(PVI_sheet.cells.last_cell.row)).end('up').row
                Pvi_last_column = PVI_sheet.range('A1').end('right').last_cell.column
                Pvi_last_column_letter=num_to_col_letters(Pvi_last_column)

                PVI_sheet.api.Range(f"A1:{Pvi_last_column_letter}{Pvi_last_row}").SpecialCells(12).Select()

                wb.app.selection.copy()

                time.sleep(1)
                wb.sheets.add(f"{key} PVI {file_month}",after=Open_gr_sheet)
                time.sleep(1)

                CHS_PVI_sheet = wb.sheets[f"{key} PVI {file_month}"]
                CHS_PVI_sheet.range(f"A1").paste()
                CHS_PVI_sheet.autofit()
                CHS_PVI_sheet.api.Cells.FormatConditions.Delete()
                PVI_sheet.activate()
                PVI_sheet.api.AutoFilterMode=False
                wb.app.api.CutCopyMode=False
                PVI_sheet.api.Cells.FormatConditions.Delete()
                Pvi_last_row = PVI_sheet.range(f'A'+ str(PVI_sheet.cells.last_cell.row)).end('up').row
                PVI_sheet.range(f'{Pvi_Links_letter_column}{Pvi_last_row+5}').expand('down').delete()

                CHS_PVI_sheet.activate()
                CHS_PVI_last_column = CHS_PVI_sheet.range('A1').end('right').last_cell.column
                CHS_PVI_last_letter_column = num_to_col_letters(CHS_PVI_last_column)
                
                CHS_PVI_column_list = CHS_PVI_sheet.range("A1").expand('right').value
                Terminal_Links_no_column=CHS_PVI_column_list.index('Terminal ')+1
                Terminal_Links_letter_column = num_to_col_letters(Terminal_Links_no_column)
                CHS_PVI_last_row = CHS_PVI_sheet.range(f'A'+ str(CHS_PVI_sheet.cells.last_cell.row)).end('up').row
                Terminal_column_value = CHS_PVI_sheet.range(f"{Terminal_Links_letter_column}2:{Terminal_Links_letter_column}{CHS_PVI_last_row}").value


                buy_sheet = wb.sheets['Buy']
                buy_sheet.activate()
                # buy_sheet.api.Range(f"A:A").Select()
                buy_sheet_last_row = buy_sheet.range(f'A'+ str(buy_sheet.cells.last_cell.row)).end('up').row
                buy_sheet.api.Range(f"A1:A{buy_sheet_last_row}").AutoFilter(Field:=1, Criteria1:=[value])
                buy_sheet_last_column = buy_sheet.range('A1').end('right').end('right').last_cell.column
                buy_sheet_last_column_letter_column = num_to_col_letters(buy_sheet_last_column)
                buy_sheet_column_list = buy_sheet.range(f"A1:{buy_sheet_last_column_letter_column}1").value
                buy_sheet_last_row = buy_sheet.range(f'A'+ str(buy_sheet.cells.last_cell.row)).end('up').row
                Purchasep_no_column=buy_sheet_column_list.index('Purchase Price')+1
                Purchasep_letter_column = num_to_col_letters(Purchasep_no_column)


                purchase_price = buy_sheet.api.Range(f"{Purchasep_letter_column}{buy_sheet_last_row}").Value
                buy_sheet.api.AutoFilterMode=False


                clist=["Voucher","Product Name","Bill No","Date","Vendor Inv. Dt.","BOLNumber","Terminal ","Account","Gross Qty","Net Qty","Billed Qty","Debit Amount"]	

                df = CHS_PVI_sheet.range(f'A1:{CHS_PVI_last_letter_column }{CHS_PVI_last_row}').options(pd.DataFrame, chunksize=10_000).value
                df = df.reset_index()
                df = df[clist]
                df = df.rename(columns={"Debit Amount":"Prov Amt"})

                a=purchase_price.replace(' ','_')
                final_price = Pricing_index_dict[a.upper().split('_MONTH')[0]+"_Ethanol"]-float(a.upper().split('-')[-1])

                filters = list(set(Terminal_column_value))

                for filter in filters:
                    temp_df = df[(df['Terminal ']==filter)]
                    temp_df.insert(loc = len(temp_df.columns)-1,column = 'Prov Price',value = round(temp_df['Prov Amt']/temp_df['Billed Qty'],5))
                    temp_df['Final Price'] = final_price
                    temp_df['Final Amt'] = round(temp_df['Final Price']*temp_df['Billed Qty'],2)
                    temp_df['Difference'] = round(temp_df['Final Amt'] - temp_df['Prov Amt'],2)
                    temp_df["Gross Qty"] = temp_df["Gross Qty"].astype(int)
                    temp_df["Net Qty"] = temp_df["Net Qty"].astype(int)
                    temp_df["Billed Qty"] = temp_df["Billed Qty"].astype(int)


                    # v['Vendor'] = [key]
                    # em_df['Location'] = filter
                    # if not lastrow:
                    try:
                        wb.sheets.add(f"{key}",after=Open_gr_sheet)
                        time.sleep(1)
                        CHS_sheet = wb.sheets[f"{key}"]
                    except:
                        CHS_sheet = wb.sheets[f"{key}"]

                    current_company_last_row = CHS_sheet.range(f'L'+ str(CHS_sheet.cells.last_cell.row)).end('up').row
                    initial_row=1
                    if initial_row!=current_company_last_row:
                        initial_row = current_company_last_row+2

                    CHS_sheet.range(f'B{initial_row}').options(index = False).value = temp_df 
                    CHS_sheet.autofit()
                    CHS_sheet.api.Range(f"{initial_row}:{initial_row}").Font.Bold = True
                    # else:
                    # CHS_sheet.range(f'B{initial_row+1}').options(index = False).value = temp_df 
                    # CHS_sheet.autofit()
                    # CHS_sheet.api.Range(f"{initial_row+1}:{initial_row+1}").Font.Bold = True 

                    # if not lastrow:
                    t_last_row = CHS_sheet.range(f'B'+ str(CHS_sheet.cells.last_cell.row)).end('up').row
                    # else:
                    #     t_last_row = lastrow+ 1
                    CHS_sheet.api.Range(f"L{t_last_row+2}").Value = f'=SUM(L{initial_row+1}:L{t_last_row})'
                    Q_amt = CHS_sheet.api.Range(f"L{t_last_row+2}").Value
                    CHS_sheet.api.Range(f"Q{t_last_row+2}").Value = f'=SUM(Q{initial_row+1}:Q{t_last_row})'
                    diff_amt = CHS_sheet.api.Range(f"Q{t_last_row+2}").Value

                    # em_df['Diff'] = round(em_df['Amount']/em_df['Qty'],4)
                    
                    # CHS_sheet.api.Range(f"L{t_last_row+2}").Value = f"Sds"
                    insert_all_borders(cellrange=f"L{t_last_row+2}",working_sheet=CHS_sheet,working_workbook=wb)
                    insert_all_borders(cellrange=f"Q{t_last_row+2}",working_sheet=CHS_sheet,working_workbook=wb)

                    CHS_sheet.api.Range(f"L{t_last_row+2}").Font.Bold = True
                    CHS_sheet.api.Range(f"Q{t_last_row+2}").Font.Bold = True

                    CHS_sheet.api.Range(f"{initial_row+1}:{initial_row+1}").Insert(Shift:=win32c.Direction.xlDown)
                    CHS_sheet.api.Range(f"{initial_row+1}:{initial_row+1}").Insert(Shift:=win32c.Direction.xlDown)
                    CHS_sheet.api.Range(f"{initial_row+1}:{initial_row+1}").Insert(Shift:=win32c.Direction.xlDown)
                    CHS_sheet.api.Range(f"B{initial_row+2}").Value = f"PO# {value}"
                    CHS_sheet.api.Range(f"I{initial_row+2}").Value = purchase_price
                    p_ters = purchase_price.split("-")[0].strip()
                    CHS_sheet.api.Range(f"O{initial_row+2}").Value = final_price
                    CHS_sheet.autofit()
                    em_df = em_df.append({'Vendor':key,'Location':filter,'Qty':Q_amt,'Amount':diff_amt,'Diff':round(diff_amt/Q_amt,4),'Pricing Terms':p_ters},ignore_index=True)
            print("done")
            wb.sheets.add(f"Summary",after=Open_gr_sheet)
            time.sleep(1)
            Summary_sheet = wb.sheets[f"Summary"]
            Summary_sheet.activate()
            if len(em_df)>0:
                Summary_sheet.range(f'B3').options(index = False).value = em_df  
                Summary_sheet.api.Range(f"3:3").Font.Bold = True
                Summary_sheet.api.Range(f"3:3").HorizontalAlignment = -4108
                Summary_sheet.autofit()
                s_last_row = Summary_sheet.range(f'D'+ str(CHS_sheet.cells.last_cell.row)).end('up').row

                Summary_sheet.api.Range(f"D{s_last_row+2}").Value = f'=SUM(D4:D{s_last_row})'
                Summary_sheet.api.Range(f"E{s_last_row+2}").Value = f'=SUM(E4:E{s_last_row})'

                insert_all_borders(cellrange=f"D{s_last_row+2}",working_sheet=Summary_sheet,working_workbook=wb)
                insert_all_borders(cellrange=f"E{s_last_row+2}",working_sheet=Summary_sheet,working_workbook=wb)

                Summary_sheet.api.Range(f"D{s_last_row+2}").Font.Bold = True
                Summary_sheet.api.Range(f"E{s_last_row+2}").Font.Bold = True
                Summary_sheet.autofit()
                Summary_sheet.api.Columns.ColumnWidth = 23
            else:
                print("em_df not created or trueup dont exist") 
            filename= rackOutput+"\\"+f"Rack AP Data {file_month} {file_year}.xlsx"
            wb.save(rackOutput+"\\"+f"Rack AP Data {file_month} {file_year}.xlsx")
        return filename
    except Exception as e:
        raise e

    finally:
        try:
            wb.app.kill()
        except:
            pass

if __name__ == "__main__":
    try:
        JOBNAME = 'AP_RACK_TRUEUP_AUTOMATION'
        receiver_email ='yashn.jain@biourja.com, devina.ligga@bioruja.com, imam.khan@biourja.com, priyanshi.jhawar@biourja.com, \
                            ravinder.rathore@biourja.com, rahul.gupta@biourja.com, arvind.patidar@biourja.com, ankur.modi@biourja.com'
        # receiver_email = 'yashn.jain@biourja.com'
        today_date = date.today()
        # today_date = datetime.strptime("12-08-2022", "%d-%m-%Y").date()
        prev_month_last_date = today_date.replace(day=1) -timedelta(days=1)
        prev_month_year = datetime.strftime(prev_month_last_date, "%m.%y")
        prev_month_year2 = datetime.strftime(prev_month_last_date, "%B %Y").upper()
        j_loc = r'J:\India\Trueup\TrueupAutomation\AP_Rack_TrueUp'
        logfile = os.getcwd()+'\\logs\\' + JOBNAME+str(today_date)+'.txt'
        trueup_file = r'J:\India\Trueup\TrueupAutomation\AP_Rack_TrueUp\Rack PO details'
        # rackInput = os.getcwd()+f"\\Input"
        rackInput = j_loc+f"\\Input"
        

        # j_loc = r"J:\India\Trueup\TrueupAutomation"
        # bulkInput = os.getcwd()+"\\Input\\bulkRaw"
        # priceInput = os.getcwd()+f"\\Prices"
        priceInput = j_loc+f"\\Prices"
        # bulkOutput = os.getcwd()+"\\Output\\Bulk_AR"
        # bulkOutput = os.getcwd()+"\\Output"
        rackOutput = j_loc+"\\Output"

        # input_mapping = j_loc+f"\\Prices\\{prev_month_year}Prices.xlsx"

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)
        logging.info("Starting AP_RACK_TRUEUP_AUTOMATION")
        filename = rackTrueup(priceInput,rackInput,trueup_file,rackOutput)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS -{JOBNAME}',mail_body = f'{JOBNAME} Completed Successfully, Attached logs',attachment_location = logfile)
    except Exception as e:
        logging.exception(e)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{JOBNAME}',mail_body = f'{JOBNAME} failed, Attached logs',attachment_location = logfile)
        sys.exit(-1)
    

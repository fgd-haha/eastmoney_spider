import xlsxwriter as xlsxwriter
def save_to_excel(output_dir, sorted_dict):
    wb_buy = xlsxwriter.Workbook(f'{output_dir}/buy.xlsx')
    ws_buy = wb_buy.add_worksheet()

    for index, code in enumerate(sorted_dict[0:20]):
        ws_buy.write(index, 0, code[1]["SName"])
        ws_buy.write(index, 1, code[0])
        ws_buy.write(index, 2, code[1]["ShareSZ_Chg_One"])
    wb_buy.close()

    wb_sell = xlsxwriter.Workbook(f'{output_dir}/sell.xlsx')
    ws_sell = wb_sell.add_worksheet()
    for index, code in enumerate(sorted_dict[-1:-21:-1]):
        ws_sell.write(index, 0, code[1]["SName"])
        ws_sell.write(index, 1, code[0])
        ws_sell.write(index, 2, code[1]["ShareSZ_Chg_One"])
    wb_sell.close()


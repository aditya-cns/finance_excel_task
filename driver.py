"""
Driver file for the program - transpose_raw_excel_file.py
"""

from transpose_raw_excel_file import SaasRevenueTransposer
import openpyxl

wb = openpyxl.load_workbook('Saas Revenue.xlsx', data_only=True)
file_name = 'Saas Revenue.xlsx'
saas_transposer = SaasRevenueTransposer(wb, file_name)
saas_transposer.create_new_excel(file_name="Saas Revenue Transposed.xlsx")

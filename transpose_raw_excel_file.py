"""
transpose_raw_excel_file converts a raw excel file into a transposed excel file.  
"""
import pandas as pd


class SaasRevenueTransposer:
    """SaasRevenueTransposer transposes the raw excel file into a transposed
    excel file where date and prices are transposed according to the customer.
    """
    def __init__(self, wb, file_name):
        self.wb = wb
        self.sheet = wb.active
        self.file_name = file_name
        self.dates = self.sheet['g1':'z1']
        self.raw_date_list = []
        self.date_list = []
        self.prices = {}
        self.date_price_dict = {}
        self.date_price_df = pd.DataFrame()
        self.customer_details_df = pd.DataFrame()
        self.rows_count = self.sheet.max_row - 1
        
        self.set_date_list()
        self.set_prices_list()
        self.itemize_date_price_df()
        self.itemize_customer_details_df()
    
    def create_new_excel(self, file_name):
        """Create a new excel file with the transposed data"""
        output_df = self.customer_details_df.merge(
            self.date_price_df,
            left_index=True, 
            right_index=True, 
            how='left'
        )
        output_df.to_excel(file_name, index=False)

    
    def itemize_customer_details_df(self):
        """Itemize the customer details into a dataframe"""
        df1 = pd.read_excel(self.file_name)
        df3 = df1.iloc[: , 0:6]
        total_dates = 20
        output_df = pd.concat([df3] * total_dates, ignore_index=True)
        self.customer_details_df = output_df
        
    
    def itemize_date_price_df(self):
        """Itemize the date and price into a dataframe"""
        pd_df_dict = {}
        date_list_expanded = []
        price_list_expanded = []
        for date_str in self.date_list:
            for _ in range(self.rows_count):
                date_list_expanded.append(date_str)
        
        for price in self.date_price_dict.values():
            for indv_price in price:
                price_list_expanded.append(indv_price)
        pd_df_dict['Date'] = date_list_expanded
        pd_df_dict['Price'] = price_list_expanded
        self.date_price_df = pd.DataFrame.from_dict(pd_df_dict)

    
    def set_prices_list(self):
        """Get the list of prices from the excel file"""
        df = pd.read_excel(self.file_name)
        df_date_price = df.iloc[: , 6:]
        self.date_price_dict = df_date_price.to_dict("list")
    
    def set_date_list(self):
        """Get the list of dates from the excel file"""
        for cell_values in self.dates:
            for date_obj in cell_values:
                self.raw_date_list.append(date_obj.value)
                self.date_list.append(str(date_obj.value.date()))

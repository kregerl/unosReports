#
import sys
import numpy as np
import pandas as pd



def group(file_in, file_out):
    frame = pd.read_csv(file_in)
    
    
    is_refund = frame['financial_status'] == 'refunded'

    is_reship = frame['net_sales'] == 0

    # Looking for the ones that are NOT from the same shipping/billing code and region
    frame['gift_order'] = np.where((frame['shipping_postal_code'] == frame['billing_postal_code']) & (frame['shipping_region'] == frame['billing_region']), True, False)
    is_gift_order = frame['gift_order'] == False

    results = {
    'Refunds': decrementing_pivot_table_creator(frame[is_refund].rename(
        columns={'shipping_region': 'State', 'financial_status' : 'Number of refunds'}), 'State', 'Number of refunds', 'count'),

    'Reships': decrementing_pivot_table_creator(frame[is_reship].rename(
        columns={'shipping_region': 'State', 'net_sales' : 'Number of reships'}), 'State', 'Number of reships', 'count'),

    'Gift Shipments': decrementing_pivot_table_creator(frame[is_gift_order].rename(
        columns={'shipping_region': 'State', 'gift_order' : 'Number of gift orders'}), 'State', 'Number of gift orders', 'count')
    }

    save_sheets_to_workbook(results, file_out)



def decrementing_pivot_table_creator(filtered_frame, index, values, func):
    frame = filtered_frame.pivot_table(index = [index], values = [values], aggfunc = func, margins=True, margins_name='Total')
    tot_frame = pd.DataFrame(frame)[len(frame[values]) -1:].copy(deep = True)
    frame = frame.drop(['Total'], axis = 0).sort_values(by = values, ascending=False)
    frame = frame.append(tot_frame)
    return frame

def save_sheets_to_workbook(results, file):
    with pd.ExcelWriter(file) as writer:
        for key in results:
            results[key].to_excel(writer, sheet_name=key)
    
if __name__ == '__main__':
    path_in = "values.csv"
    path_out = 'final.xlsx'
    group(path_in, path_out)

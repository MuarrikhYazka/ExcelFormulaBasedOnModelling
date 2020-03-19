import xlwings as xw
import statistics
from scipy import stats
import pickle


def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def deploy(cashflow_6, cashflow_5, cashflow_4, cashflow_3, cashflow_2, cashflow_1, transaction_6, transaction_5, transaction_4, transaction_3, transaction_2, transaction_1):
    list_bulan=[1,2,3,4,5,6]
    cashflow=[cashflow_6, cashflow_5, cashflow_4, cashflow_3, cashflow_2, cashflow_1]
    transaction=[transaction_6, transaction_5, transaction_4, transaction_3, transaction_2, transaction_1]
    stdev_c=statistics.stdev([cashflow_1, cashflow_2, cashflow_3, cashflow_4, cashflow_5, cashflow_6])
    slope_c, intercept_c, r_value_c, p_value_c, std_err_c=stats.linregress(cashflow, list_bulan)
    stdev_t=statistics.stdev([transaction_1, transaction_2, transaction_3, transaction_4, transaction_5, transaction_6])
    slope_t, intercept_t, r_value_t, p_value_t, std_err_t=stats.linregress(transaction, list_bulan)


    with open('D:\Task1510\saved_model.pkl', 'rb') as f:
        model = pickle.load(f)

    input=[[stdev_t,stdev_c,slope_t,slope_c]]
    probability=model.predict(input)
    #probability=cashflow_6+cashflow_5

    
    return probability



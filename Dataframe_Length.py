import pandas

def count(dataframe):

    ph_bc = len(dataframe[0].index)
    ph_pip = len(dataframe[1].index)
    ph_new = len(dataframe[2].index)
    ph_error = len(dataframe[3].index)

    return str(ph_bc), str(ph_pip), str(ph_new), str(ph_error)
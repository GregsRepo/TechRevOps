from pandas import ExcelWriter

###########################################################################
def excel(DF, DIR, EXCEL):

    DF.to_excel(DIR + EXCEL, index=False)

###########################################################################
def ph_to_excel(ph_status_df, ph_aging_df, DIR):

    # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    with ExcelWriter(DIR + 'Provisioning_Report.xlsx') as writer:

        ph_status_df[0].to_excel(writer, sheet_name='BC Status', index=False)
        ph_status_df[1].to_excel(writer, sheet_name='PIP Status', index=False)
        ph_status_df[2].to_excel(writer, sheet_name='New Status', index=False)
        ph_status_df[3].to_excel(writer, sheet_name='PE Status', index=False)
        ph_status_df[4].to_excel(writer, sheet_name='Status Review', index=False)

        ph_aging_df[0].to_excel(writer, sheet_name='BC Aging', index=False)
        ph_aging_df[1].to_excel(writer, sheet_name='PIP Aging', index=False)
        ph_aging_df[2].to_excel(writer, sheet_name='New Aging', index=False)
        ph_aging_df[3].to_excel(writer, sheet_name='PE Aging', index=False)
        ph_aging_df[4].to_excel(writer, sheet_name='Aging Review', index=False)

###########################################################################
def zaci_to_excel(dx, dme, credit_hold, DIR):

    # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    with ExcelWriter(DIR + 'ZACI_Report.xlsx') as writer:
        dx.to_excel(writer, sheet_name='DX', index=False)
        dme.to_excel(writer, sheet_name='DME', index=False)
    
    with ExcelWriter(DIR + 'Credit_Hold.xlsx') as writer:
        credit_hold.to_excel(writer, sheet_name='Credit Hold', index=False)
        
###########################################################################

 

        
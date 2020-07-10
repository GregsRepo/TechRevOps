from pandas import ExcelWriter

###########################################################################
def excel(DF, DIR, EXCEL):

    DF.to_excel(DIR + EXCEL, index=False)

###########################################################################
def ph_to_excel(ph_status_df, ph_aging_df, DIR):

    # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    with ExcelWriter(DIR + 'Provisioning Report.xlsx') as writer:

        ph_status_df[0].to_excel(writer, sheet_name='Status New', index=False)
        ph_status_df[1].to_excel(writer, sheet_name='Status Complete', index=False)
        ph_status_df[2].to_excel(writer, sheet_name='Status Complete', index=False)
        ph_status_df[3].to_excel(writer, sheet_name='Status In Progress', index=False)
        ph_status_df[4].to_excel(writer, sheet_name='Status Review', index=False)

        ph_aging_df[0].to_excel(writer, sheet_name='Aging New', index=False)
        ph_aging_df[1].to_excel(writer, sheet_name='Aging Complete', index=False)
        ph_aging_df[2].to_excel(writer, sheet_name='Aging In Progress', index=False)
        ph_aging_df[3].to_excel(writer, sheet_name='Aging Prov Error', index=False)
        ph_aging_df[4].to_excel(writer, sheet_name='Aging Review', index=False)

        
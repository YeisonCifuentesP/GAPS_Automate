from automate_GAPS import getGaps
from automate_GAPS import *
from os import remove



gapsRig23,dfGaps23 = getGaps('IndependenceRig23-ALL')
gapsRig30,dfGaps30 = getGaps('IndependenceRig30-ALL')
gapsRig43,dfGaps43 = getGaps('IndependenceRig43-ALL')




name_excel='GAPS_EXPORT_SIERRACOL.xlsx'
writer = pd.ExcelWriter(name_excel, engine='xlsxwriter')
dfGaps23.to_excel(writer,sheet_name='RIG_23', index = False)
dfGaps30.to_excel(writer,sheet_name='RIG_30', index = False)
dfGaps43.to_excel(writer,sheet_name='RIG_43', index = False)
writer.save()
writer.close()
send_email(gapsRig23, gapsRig30, gapsRig43,name_excel)
remove(name_excel)
        
print("Correo enviado con exito")


    
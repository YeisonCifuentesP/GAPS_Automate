from automate_GAPS import *
import os
import time
import sys


path='D:\Independence\Sistema\Desktop\GAPS_EXPORT_SIERRACOL.xlsx'
try:
    
    gapsRig23,dfGaps23 = getGaps('IndependenceRig23-ALL')
    gapsRig30,dfGaps30 = getGaps('IndependenceRig30-ALL')
    gapsRig43,dfGaps43 = getGaps('IndependenceRig43-ALL')

    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    dfGaps23.to_excel(writer,sheet_name='RIG_23', index = False)
    dfGaps30.to_excel(writer,sheet_name='RIG_30', index = False)
    dfGaps43.to_excel(writer,sheet_name='RIG_43', index = False)   
    writer.close()
    send_email(gapsRig23, gapsRig30, gapsRig43,path)
    os.remove(path)            
    print("Correo enviado con exito")
except BaseException as ex:        
    ex_type, ex_value, ex_traceback = sys.exc_info()
    message='Opps ha ocurrido un error: ' + str(ex_type)+str(ex_value) 
    print(message)

time.sleep(10)
    

    
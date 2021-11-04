from automate_GAPS import getGaps
from automate_GAPS import send_email




gapsRig23 = getGaps('IndependenceRig23-ALL')
gapsRig30 = getGaps('IndependenceRig30-ALL')
gapsRig43 = getGaps('IndependenceRig43-ALL')
send_email(gapsRig23, gapsRig30, gapsRig43)
print("Correo enviado con exito")



    
from classes.funciones import export_excel_infoallboxes, export_excel_infofilteredboxes, export2excel, ruta_corrector

ruta=r'C:/Users/Adrian Moreno/Downloads/310_02_Deliverable_STW+CajasUnion.ifc'

ruta_corregida=ruta_corrector(ruta)
# ruta = "C:/Users/Adrian Moreno/Downloads/310_02_Deliverable_STW+CajasUnion.ifc"

input_usuario=input("""Qué opción deseas exportar a excel?
    a) Todas las cajas del modelo
    b) La info de los parent joint, excluyendo repetidos
    c) La info de los parent joint, excluyendo repetidos, e incluyendo número de openings que generan HoldDown
    """)

print("Has elegido la opción: "+"'"+input_usuario+"'")


if input_usuario=="a":
    export_excel_infoallboxes(ruta_corregida)
    

if input_usuario=="b":
    export_excel_infofilteredboxes(ruta_corregida)
    

if input_usuario=="c":
    export2excel(ruta_corregida)


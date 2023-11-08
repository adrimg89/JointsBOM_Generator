from classes.funciones import export_excel_infoallboxes, export_excel_infofilteredboxes, export2excel

# ruta='C:/011H/011h - 02 - INITIATIVES/02 - DIGITAL COMPONENTS & SYSTEMS/05 - JOINTS/03 Documentos WIP/Joints 2/2201_DZ/231107_Cajas DZ. Exportación de materiales/20231031_2201-01_011_04_MOD_ARC_V01_16.ifc'

ruta = "C:/Users/Adrian Moreno/Downloads/310_02_Deliverable_STW+CajasUnion.ifc"

input_usuario=input("""Qué opción deseas exportar a excel?
    a) Todas las cajas del modelo
    b) La info de los parent joint, excluyendo repetidos
    c) La info de los parent joint, excluyendo repetidos, e incluyendo número de openings que generan HoldDown
    """)

print("Has elegido la opción: "+"'"+input_usuario+"'")


if input_usuario=="a":
    export_excel_infoallboxes(ruta)
    

if input_usuario=="b":
    export_excel_infofilteredboxes(ruta)
    

if input_usuario=="c":
    export2excel(ruta)


from classes.funciones import export_excel_infoallboxes, export_excel_infofilteredboxes, export_excel_infofilteredboxes_withnrofbalconies


ruta = "C:/Users/Adrian Moreno/Downloads/7201_RAT TGL - Distrito Z 7201-01_011_06_MOD_ARC_WIP #1.ifc"

input_usuario=input("""Qué opción deseas?
    a) exportar un excel con la info de todas las cajas del modelo
    b) exportar un excel con la info de cada ParentJoint filtrando los repetidos
    c) exportar un excel de Parent Joint únicos con la información de openings asociada
    """)


if input_usuario=="a":
    export_excel_infoallboxes(ruta)
    

if input_usuario=="b":
    export_excel_infofilteredboxes(ruta)
    

if input_usuario=="c":
    export_excel_infofilteredboxes_withnrofbalconies(ruta)
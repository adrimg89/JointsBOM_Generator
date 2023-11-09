from classes.funciones import getmodeledconnections
from classes.funciones_2 import rl_cgtype_ctype

# ruta = 'C:/011H/011h - 02 - INITIATIVES/02 - DIGITAL COMPONENTS & SYSTEMS/05 - JOINTS/03 Documentos WIP/WIP RAT/Revisión modelo RAT cajas/230901_Modelo final con todas las mejoras/230901_7201-01_011_06_MOD_STC&JOINTS_V03.01_WIP.ifc'
ruta = 'C:/011H/011h - 02 - INITIATIVES/02 - DIGITAL COMPONENTS & SYSTEMS/05 - JOINTS/03 Documentos WIP/Modelos test/Unitarios Sistemas y Componentes Digitales 7000-00_4_MOD_RVT22_JOINTS #5.ifc'

# for i in getmodeledconnections(ruta):print(i)

# connections=rl_cgtype_ctype()

# for i in connections:
#     if i['is_modeled'][0]=='Yes':print(i)

modeledconnections=getmodeledconnections(ruta)

def contar_repeticiones(connectiontypesmodeled):
    recuento = {}
    
    for item in connectiontypesmodeled:
        parent_joint_id = item['JS_ParentJointInstanceID']
        connection_type_id = item['JS_ConnectionTypeID']
        
        if parent_joint_id in recuento:
            if connection_type_id in recuento[parent_joint_id]:
                recuento[parent_joint_id][connection_type_id] += 1
            else:
                recuento[parent_joint_id][connection_type_id] = 1
        else:
            recuento[parent_joint_id] = {connection_type_id: 1}
    
    resultado = []
    
    for parent_joint_id, connection_types in recuento.items():
        for connection_type_id, count in connection_types.items():
            resultado.append({
                'ParentJoint_id': parent_joint_id,
                'Connectiontype_id': connection_type_id,
                'nr_units': count
            })
    
    return resultado

conexiones_modeladas=contar_repeticiones(modeledconnections)

for i in conexiones_modeladas:print(i)


from classes.funciones import get_instance_complete_boxes,bomlines_joints,bomlines_inferredconnections,bomlines_modeledconnections,connectiongroup_records, airtable_conection, connectiontype_records, instanciar_cgtype, instanciar_ctype, joints3playground_connect, rl_cgtype_ctype_records, ctypetocgtype,connectionlayers_records, instanciar_clayer,clayertoctype, materials_records, instanciar_materials, materialtoclayer,layercost,connectioncost, materialgroups_records,matlayers_records,instanciar_matgroups,instanciar_matlayers,matlayertomatgroup,materialtomatlayer,joints2_records,instanciar_joints,joints2layers_records,instanciar_jointslayers,jointlayertojoint,materialtojointlayer,instanciarIFC,get_allboxes,filter_boxes,instanciarboxes,nropenings_to_boxes,get_modeledconnections,instanciarconexionesmodeladas

from passwords import *

# joints3playground_connect=joints3playground_connect

ruta=input(r'Introduce la ruta del archivo IFC: ')

boxes_objects,herrajes_objects=get_instance_complete_boxes(ruta)

joints_bom_lines=bomlines_joints(boxes_objects)

# modeledconnections_bom_lines=bomlines_modeledconnections(boxes_objects,herrajes_objects)
inferredconnections_bom_lines=bomlines_inferredconnections(boxes_objects)

# for i in boxes_objects:
#     print(i.id,'Joint Type:',i.joint_type, i.corematgroups,i.q1matgroups,i.q2matgroups,i.q3matgroups,i.q4matgroups,i.connectiongroup_type)

# for box in boxes_objects:
#     joint_layers=[]
#     connection_layers=[]
    
#     if box.id=='PJ.V-0001':
#         joint=box.joint_type        
#         if joint != '':
#             for i in joint.matlayers:
#                 joint_layers.append(i)
            
#     for layer in joint_layers:
#         print(layer.performance,layer.calcformula,layer.fase,layer.material.sku,layer.material.description,layer.material.unit,box.id,layer.cost)
        
# for line in inferredconnections_bom_lines:
#     if line.parent=='PJ.H-0104':
#         print(line.performance,line.calcformula,line.ctype,line.sku,line.materialcost,line.unit,line.parent,line.cgtype,line.floor,'x',line.quantity,'Cost:',line.layercost)


def jointline_to_dictlist(joints_bom_lines):
    lista=[]
    for line in joints_bom_lines:
        parent=line.parent
        performance=line.performance
        calcformula=line.calcformula
        fase=line.fase
        description=line.description
        jointtype=line.jointtype.id
        sku=line.sku
        materialcost=line.materialcost
        unit=line.unit        
        quantity=line.quantity
        layercost=line.layercost        
        floor=line.floor               
        
        diccionario={'Performance':performance,
                     'Calculation Formula':calcformula,
                     'Fase':fase,
                     'JointTypeid/MaterialGroup':jointtype,
                     'SKU':sku,
                     'Estimated Material Cost':materialcost,
                     'Description':description,
                     'Unit':unit,
                     'Parent_ID':parent,
                     'Quantity':quantity,
                     'Layer Cost':layercost,
                     'Floor':floor}
        lista.append(diccionario)
    return lista

def inferredconnectionline_to_dictlist(inferredconnections_bom_lines):
    lista=[]
    for line in inferredconnections_bom_lines:
        parent=line.parent
        performance=line.performance
        calcformula=line.calcformula
        fase=line.fase
        description=line.description
        ctype=line.ctype        
        sku=line.sku
        materialcost=line.materialcost
        unit=line.unit 
        ismodeled=line.ismodeled       
        quantity=line.quantity
        layercost=line.layercost 
        cgtype=line.cgtype       
        floor=line.floor  
        materialtype=line.materialtype             
        
        diccionario={'Performance':performance,
                     'Calculation Formula':calcformula,
                     'Fase':fase,
                     'Description':description,
                     'Connection Type':ctype,
                     'SKU':sku,
                     'Estimated Material Cost':materialcost,
                     'Unit':unit,
                     'Is Modeled?':ismodeled,
                     'Parent_ID':parent,
                     'Quantity':quantity,
                     'Layer Cost':layercost,
                     'Connectiongroup type':cgtype,
                     'Floor':floor}
        lista.append(diccionario)
    return lista


import pandas as pd
import os



def export2excel(ruta,lista_de_diccionarios):    
    # Obtener el nombre del archivo IFC sin la extensión
    nombre_archivo = os.path.splitext(os.path.basename(ruta))[0]
    
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(lista_de_diccionarios)
    
    # Obtener la ruta del directorio donde se encuentra el archivo IFC
    directorio_ifc = os.path.dirname(ruta)
    
    # Generar el nombre del archivo Excel con el sufijo 'filteredboxesandnrbalconies'
    excel_file_name = f"{nombre_archivo}_prueba.xlsx"
    
    # Combinar la ruta del directorio del archivo IFC con el nombre del archivo Excel
    ruta_excel = os.path.join(directorio_ifc, excel_file_name)
    
    # Exportar el DataFrame a un archivo de Excel en la misma ruta que el archivo IFC
    df.to_excel(ruta_excel, index=False)
    
    print(f'Exportación a {excel_file_name} completada en la ruta: {directorio_ifc}')
    

lista_de_diccionarios=inferredconnectionline_to_dictlist(inferredconnections_bom_lines)

export2excel(ruta,lista_de_diccionarios)   
    

# ifc_object=instanciarIFC(ruta)

# allboxes=get_allboxes(ifc_object)

# filtered_boxes=filter_boxes(allboxes)

# boxes_objects=instanciarboxes(filtered_boxes)

# nropenings_to_boxes(ifc_object,boxes_objects)

# for i in boxes_objects:
#     if i.nrbalconies!=0:
#         print(i.id,i.nrbalconies)

      

# connectiongroups=connectiongroup_records() #{'Description': 'NINO15080+ 31 x LBA460 c/1000mm, 2 x (LBV100800 + 90 x LBA460) en extremos, 2 x (LBV100800 + 90 x LBA460) en puertas interiores (flejes por el interior)', 'param_anglecadence': '800 mm', 'param_endHD': 'LBV100800', 'param_balconyHD': 'LBV100800', 'cgtype_id': 'CG_0253', 'RL_ConnectionGroupType_ConnectionType_api': 'CG_0253+H_T2-0006-x2, CG_0253+H_T2-0006-x2, CG_0253+H_T3-0060-c/800', 'api_ConnectionGroup_Class': 'ESC.FLEJ'}
# connectiontypes=connectiontype_records() #{'description': 'NINO15080 + 31 x LBA460', 'is_modeled': 'Yes', 'connection_type_id': 'H_T3-0060', 'RL_ConnectionGroupType_Connectiontype_api': 'CG_0250+H_T3-0060-c/1000, CG_0251+H_T3-0060-c/1000, CG_0252+H_T3-0060-c/800, CG_0253+H_T3-0060-c/800'}
# relations=rl_cgtype_ctype_records() #{'Performance': 4, 'Calculation Formula': 'Length * performance', 'RL_cgtype_ctype': 'CG_0237+H_C1-0044-c/250', 'connectiongroup_type_id': 'CG_0237', 'connection_type': 'H_C1-0044'}
# clayers=connectionlayers_records()
# materials=materials_records()
# materialgroups=materialgroups_records()
# matlayers=matlayers_records()
# joints=joints2_records()
# jointlayers=joints2layers_records()

# connectiongroup_objects=instanciar_cgtype(connectiongroups)
# connectiontype_objects=instanciar_ctype(connectiontypes)
# connectionlayer_objects=instanciar_clayer(clayers)
# material_objects=instanciar_materials(materials)
# matgroup_objects=instanciar_matgroups(materialgroups)
# matlayer_objects=instanciar_matlayers(matlayers)
# joints_objects=instanciar_joints(joints)
# jointlayers_objects=instanciar_jointslayers(jointlayers)

# asignarctypesacgtypes=ctypetocgtype(connectiongroup_objects,relations,connectiontype_objects)
# asignarclayersactypes=clayertoctype(connectiontype_objects,connectionlayer_objects)
# asignarmaterialsaclayers=materialtoclayer(connectionlayer_objects,material_objects)
# asignarmatlayersamatgroups=matlayertomatgroup(matgroup_objects,matlayer_objects)
# asignarmaterialesamatlayers=materialtomatlayer(matlayer_objects,material_objects)
# asignarjointlayersajoints=jointlayertojoint(joints_objects,jointlayers_objects)
# asignarmaterialesajlayers=materialtojointlayer(jointlayers_objects,material_objects)

# allLayers=jointlayers_objects+matlayer_objects+connectionlayer_objects

# asignarlayercost=layercost(allLayers)
# asignarconnectioncost=connectioncost(connectiontype_objects)

# connectiontypes=get_connectiontype() # {'connection_type_id': 'H_T3-0060', 'description': ['NINO15080 + 31 x LBA460'], 'formula': 'Length * performance', 'performance': 1, 'CGtype_description': None}
# clayers=get_clayers() # {'connectiontype_id': 'H_T9-0031', 'material': 'MBRA0801', 'description': None, 'performance': 1, 'formula': 'Fix value', 'fase': 'Onsite for Assembly'}

# for i in connectiongroups: print(i)

# ruta = r"C:\Users\Adrian Moreno\Downloads\DZ P2_v03.ifc"
# ruta_buena = ruta_corrector(ruta)

# boxes=getboxesfilteredwithbalconies(ruta_buena)

# objetos=[]

# for i in boxes:
#     objeto=box(i['JS_ParentJointInstanceID'],i['Box_type'], i['JS_JointType'],i['JS_JointTypeID'],i['JS_ConnectionGroupTypeID'],i['QU_Length_m'],i['Core Matgroup'], i['Q1 Matgroup'],i['Q2 Matgroup'], i['Q3 Matgroup'], i['Q4 Matgroup'],i['EI_LocalisationCodeFloor'], i['nrbalconies'])
#     objetos.append(objeto)

# for i in objetos: print(i.id)

# connectiongroup_dict={'Description': 'VGZ7280 c/150mm', 'param_screwlong': 'VGZ 7x280 mm', 'param_screwcadence': '150 mm', 'cgtype_id': 'CG_0246', 'api_ConnectionGroup_Class': 'TIR90H'}



# for i in connectiongroup_objects:
#     if i.id=='CG_0252':
#         print('')
#         print('Connectiongroup:',i.id)
#         ctypes=i.connectiontypes
#         print('')
#         print('Datos de las ConnectionTypes:')
#         for ctype in ctypes:
#             print('')
#             print('     ConnectionType_id:',ctype['connection_type_id'])
#             print('     Performance:',ctype['Performance'])
#             print('     Calculation Formula:',ctype['Calculation Formula'])
        
            
#         for ctype in ctypes:
#             print('')
#             print('Layers del Connection Type: ', ctype['connection_type'].id)
#             ctype_object=ctype['connection_type']
#             clayers=ctype_object.connection_layers
#             for clayer in clayers:
#                 print('     Material:',clayer.material)
#                 print('     Performance:',clayer.performance)
#                 print('     Calculation Fromula:',clayer.calcformula)
#                 print('')

# for i in matgroup_objects:
#     print(i.id, i.description)
#     for matlayer in i.matlayers:
#         print(matlayer.material.sku,matlayer.material.description,matlayer.performance,matlayer.calcformula,matlayer.fase)
#     print('')

# for i in joints_objects:
#     print (i.id)
#     for jlayer in i.matlayers:
#         print(jlayer.material.sku,jlayer.material.description)



# def jointchecker(joints_objects):
#     joints_objects_checked=[]
#     joints_objects_notvalid=[]
#     for i in joints_objects:
#         counter=0
#         if i.matlayers==[]:
#             counter=counter+1
#         if i.matlayers!=[]:
#             layers=i.matlayers
#             for layer in layers:
#                 if layer.material is None or layer.material.type=='Connection':
#                     counter=counter+1
#         if counter==0:
#             joints_objects_checked.append(i)
#         else:
#             joints_objects_notvalid.append(i)
#     return joints_objects_checked,joints_objects_notvalid

# checkedjoints, notvalidjoints=jointchecker(joints_objects)

# for i in checkedjoints:
#     print (i.id)
    
# for i in joints_objects:
#     if i in checkedjoints:
#         print(i.id, 'Valid')
#     elif i in notvalidjoints:
#         print(i.id, 'Not Valid')    
    
# for i in checkedjoints:
#     if i.id=='J_0108':
#         print(i.id)
#         for layer in i.matlayers:
#             print('SKU:',layer.material.sku,'Description:',layer.material.description,'Performance:',layer.performance,'Coste Layer:',layer.cost)
    

# def connectiongroupcost_calculator(connectiongroup_type,long,connectiongroup_objects):
#     for i in connectiongroup_objects:
#         if i.id==connectiongroup_type:
#             print(i.connectiontypes)
#             for connection in i.connectiontypes:
#                 ctype=connection['connection_type']
#                 for layer in ctype.connectionlayers:
#                     print(layer.material_sku,layer.material.description,'x', layer.performance)

# connectiongroupcost_calculator('CG_0001',3,connectiongroup_objects)




def boxlayers(connectiongroup_type,long, nrbalconies,connetiongroup_objects):
    layers=[]
    for cgtype in connetiongroup_objects:
        if connectiongroup_type == cgtype.id:
            for contype in cgtype.connectiontypes:
                ctype=contype['connection_type']
                ctype_performance=contype['Performance']
                ctype_calculationformula=contype['Calculation Formula']
                if ctype_calculationformula=='Fix value':
                    ctype_multiplicador=ctype_performance
                elif ctype_calculationformula=='Length * performance':
                    ctype_multiplicador=ctype_performance*long
                elif ctype_calculationformula=='Opening * performance':
                    ctype_multiplicador=ctype_performance*nrbalconies
                for clayer in ctype.connectionlayers:
                    clayer_sku=clayer.material.sku
                    material_cost=clayer.material.cost
                    clayer_description=clayer.material.description
                    clayer_unit=clayer.material.unit
                    clayer_type=clayer.material.type
                    clayer_quantity=clayer.performance*ctype_multiplicador
                    clayer_cost=clayer.cost*ctype_multiplicador                    
                    layer={'SKU':clayer_sku,
                        'Material Unitary Cost':round(material_cost,2),
                        'Description':clayer_description,
                        'Measurement Unit':clayer_unit,
                        'Type of Material':clayer_type,
                        'Quantity':clayer_quantity,
                        'Layer Cost':round(clayer_cost,2)
                        }
                    layers.append(layer)
    return layers

# box_layers=boxlayers('CG_0003',3,0,connectiongroup_objects)

# for i in box_layers: print(i)

# modeledconnections=get_modeledconnections(ifc_object)
# herrajes_objetos=instanciarconexionesmodeladas(modeledconnections,connectiontype_objects)

# for herraje in herrajes_objetos:
#     print(herraje.connectiontypeid.id,herraje.connectiontypeid.description)


import ifcopenshell
import ifcopenshell.util
import ifcopenshell.util.element
import pandas as pd
import os

def abrir_ifc(ruta):
    ifc=ifcopenshell.open(ruta)
    return ifc

def list_elements(ifc, type, pset, parameter, value):
    boxes=[]
    building_elem_proxy_collector = ifc.by_type(type)
    for i in building_elem_proxy_collector:
        i_psets=ifcopenshell.util.element.get_psets(i)
        if pset in i_psets:
            parameter_value=i_psets[pset].get(parameter, '')
            if value in parameter_value:
                boxes.append(i)
    return boxes

def boxes_info_joint(ruta):
    
    boxes_info=[]
    
    ifc=abrir_ifc(ruta)
    type='IfcBuildingElementProxy'
    pset='EI_Elements Identification'
    parameter='EI_Type'
    value='Joint'
    boxes=list_elements(ifc, type, pset, parameter, value)
    
    
    for i in boxes:
        i_psets = ifcopenshell.util.element.get_psets(i)
        JS_Joint_Specification = i_psets.get('JS_Joint Specification', {})
        QU_Quantity=i_psets.get('QU_Quantity',{})
        Pset_QuantityTakeOff=i_psets.get('Pset_QuantityTakeOff',{})
        JointTypeID = JS_Joint_Specification.get('JS_JointTypeID', '')
        cgt=JS_Joint_Specification.get('JS_ConnectionGroupTypeID', '')
        parent_joint_id = JS_Joint_Specification.get('JS_ParentJointInstanceID', '')
        r_guid = i_psets['EI_Interoperability'].get('RevitGUID', '')
        QU_Length=QU_Quantity.get('QU_Length_m', '')
        Box_type=Pset_QuantityTakeOff.get('Reference', '')
        inst_a=JS_Joint_Specification.get('JS_C01_ID', '')
        inst_b=JS_Joint_Specification.get('JS_C02_ID', '')
        parameters_info={'RevitGUID':r_guid, 'JS_ParentJointInstanceID':parent_joint_id, 'JS_JointTypeID':JointTypeID, 'JS_ConnectionGroupTypeID':cgt, 'QU_Length_m':QU_Length, 'Box_type':Box_type, 'JS_C01_ID':inst_a, 'JS_C02_ID':inst_b}
        boxes_info.append(parameters_info)
    
    return boxes_info

def get_boxesfilteredbyPJ(ruta_archivo):
    
    allboxes_info=boxes_info_joint(ruta_archivo)

    parents_únicos=set()

    allboxes_info_filtered=[]

    for info in allboxes_info:
        parent=info['JS_ParentJointInstanceID']
        #print(parent)
        
        if parent not in parents_únicos:
            parents_únicos.add(parent)
            #print(parents_únicos)
            allboxes_info_filtered.append(info)
                
    return allboxes_info_filtered

def get_balconys(ruta):
    
    balcony=[]
    ifc=abrir_ifc(ruta)
    type='IfcWindow'
    filterbytype=ifc.by_type(type)
    
    for i in filterbytype:
        i_psets=ifcopenshell.util.element.get_psets(i)
        QU_Quantity = i_psets.get('QU_Quantity', {})
        EI_Elements_Identification = i_psets.get('EI_Elements Identification', {})
        EI_OpeningType = EI_Elements_Identification.get('EI_OpeningType', '')
        EI_HostComponentInstanceID=EI_Elements_Identification.get('EI_HostComponentInstanceID', '')
        QU_Height=QU_Quantity.get('QU_Height_m', '')
        if QU_Height>1.8:
            balcony.append({'EI_OpeningType':EI_OpeningType,'EI_HostComponentInstanceID':EI_HostComponentInstanceID,'QU_Height_m':QU_Height})
    return balcony

def get_huecosdepaso(ruta):
    
    doorspaso_info=[]
    ifc=abrir_ifc(ruta)
    type='IfcWall'
    filterbytype=ifc.by_type(type)
    
    for i in filterbytype:
        i_psets=ifcopenshell.util.element.get_psets(i)
        Pset_ProductRequirements=i_psets.get('Pset_ProductRequirements',{}) 
        EI_Elements_Identification=i_psets.get('EI_Elements Identification',{})   
        EI_HostComponentInstanceID=EI_Elements_Identification.get('EI_HostComponentInstanceID','')
        EI_OpeningType=EI_Elements_Identification.get('EI_OpeningType', '') 
        Category=Pset_ProductRequirements.get('Category', '')
        Params_info={'EI_HostComponentInstanceID':EI_HostComponentInstanceID,'EI_OpeningType':EI_OpeningType, 'Category':Category}
        if Category=='Doors':
                doorspaso_info.append(Params_info)
    return doorspaso_info

def get_doors(ruta):
    doors_info=[]
    ifc=abrir_ifc(ruta)
    type='IfcDoor'
    filterbytype=ifc.by_type(type)
    
    for i in filterbytype:
        i_psets=ifcopenshell.util.element.get_psets(i)
        Pset_ProductRequirements=i_psets.get('Pset_ProductRequirements',{}) 
        EI_Elements_Identification=i_psets.get('EI_Elements Identification',{})
        EI_HostComponentInstanceID=EI_Elements_Identification.get('EI_HostComponentInstanceID','')
        EI_OpeningType=EI_Elements_Identification.get('EI_OpeningType', '') 
        Category=Pset_ProductRequirements.get('Category', '')
        Params_info={'EI_HostComponentInstanceID':EI_HostComponentInstanceID,'EI_OpeningType':EI_OpeningType, 'Category':Category}
        doors_info.append(Params_info)
    return doors_info

def get_alldoors(ruta):
    alldoors_info=get_doors(ruta)+get_huecosdepaso(ruta)
    return alldoors_info

def nrtallopenings_byinstance(ruta):
    info_ifcwindow=get_balconys(ruta)
    alldoors_info=get_alldoors(ruta)
    recuento_balconerasporinstancia={}
    
    for balcony in info_ifcwindow:
        instancia=balcony['EI_HostComponentInstanceID']
        if instancia in recuento_balconerasporinstancia:
            recuento_balconerasporinstancia[instancia] += 1
        else:
            recuento_balconerasporinstancia[instancia] = 1
            
    for door in alldoors_info:
        instancia=door['EI_HostComponentInstanceID']
        if instancia in recuento_balconerasporinstancia:
            recuento_balconerasporinstancia[instancia] += 1
        else:
            recuento_balconerasporinstancia[instancia] = 1
    
    return recuento_balconerasporinstancia

def getboxesfilteredwithbalconies(ruta):
    balcony_instances_input=nrtallopenings_byinstance(ruta)
    allboxes_info_filtered_input=get_boxesfilteredbyPJ(ruta)
    boxesandbalconys=[]
    #PJ='C_FAC-0022_7203-00.0474'
    #print(balcony_instances_input[PJ])
    for i in allboxes_info_filtered_input:
        componente=i['JS_C01_ID']
        if i['Box_type']=="H.ST_Bottom":
            if componente!='' and componente in balcony_instances_input.keys():
                balconerasdelainstancia=balcony_instances_input[componente]
                i['nrbalconies']=balconerasdelainstancia
                boxesandbalconys.append(i)
                #print(i)
            else:
                i['nrbalconies']=0
                boxesandbalconys.append(i)
        else:
            i['nrbalconies']=0
            boxesandbalconys.append(i)
    return boxesandbalconys

def export2excel(ruta, boxesinfo):
    nombrearchivo=os.path.splitext(os.path.basename(ruta))[0]
    df = pd.DataFrame(boxesinfo)
    ruta_excel = os.path.join(os.path.dirname(ruta), f'{nombrearchivo}.xlsx')
    df.to_excel(ruta_excel, index=False)
    print('Exportación a excel completada')
    
def export_excel_infoallboxes(rutaifc):
    
    datos_export=boxes_info_joint(rutaifc)
            
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(datos_export)
    # Exportar el DataFrame a un archivo de Excel
    export_path="C:/Users/Adrian Moreno/Desktop"
    excel_file_path=os.path.join(export_path, "allboxes_project.xlsx")
    df.to_excel(excel_file_path, index=False)  
    
    print("Resultados exportados a allboxes_project.xlsx")
    
def export_excel_infofilteredboxes(rutaifc):
    
    datos_export=get_boxesfilteredbyPJ(rutaifc)
            
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(datos_export)
    # Exportar el DataFrame a un archivo de Excel
    export_path="C:/Users/Adrian Moreno/Desktop"
    excel_file_path=os.path.join(export_path, "filteredboxes_project.xlsx")
    df.to_excel(excel_file_path, index=False)  
    
    print("Resultados exportados a filteredboxes_project.xlsx")
    

def getmodeledconnections(ruta):
    connections_info=[]    
    ifc=abrir_ifc(ruta)
    type='IfcBuildingElementProxy'
    pset='EI_Elements Identification'
    parameter='EI_Type'
    value='Connection'
    connections=list_elements(ifc, type, pset, parameter, value)
    for i in connections:
        i_psets = ifcopenshell.util.element.get_psets(i)        
        JS_Joint_Specification = i_psets.get('JS_Joint Specification', {})    
        EI_Interoperability=i_psets.get('EI_Interoperability', {})
        connectiontype_id = JS_Joint_Specification.get('JS_ConnectionTypeID', '')        
        parent_joint_id = JS_Joint_Specification.get('JS_ParentJointInstanceID', '')
        r_guid = EI_Interoperability.get('RevitGUID', '')      
        parameters_info={'RevitGUID':r_guid,'JS_ParentJointInstanceID': parent_joint_id, 'JS_ConnectionTypeID':connectiontype_id}
        connections_info.append(parameters_info)        
    return connections_info
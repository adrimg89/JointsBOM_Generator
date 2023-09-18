import ifcopenshell
import ifcopenshell.util
import ifcopenshell.util.element
import pandas as pd
import os

def abrir_ifc(ruta):
    ifc=ifcopenshell.open(ruta)
    return ifc

def list_boxes(ifc, type, pset, parameter, value):
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
    boxes=list_boxes(ifc, type, pset, parameter, value)
    
    
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
        parameters_info=[r_guid, parent_joint_id, JointTypeID, cgt, QU_Length, Box_type, inst_a, inst_b]
        boxes_info.append(parameters_info)
    
    return boxes_info

def export_excel_infoallboxes(rutaifc):
    
    datos_export=boxes_info_joint(rutaifc)
            
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(datos_export, columns=["r_guid", "parent_joint_id", "JointTypeID", "CGT", "QU_Length", "Box_type", "Inst A", "Inst B"])
    # Exportar el DataFrame a un archivo de Excel
    export_path="C:/Users/Adrian Moreno/Desktop"
    excel_file_path=os.path.join(export_path, "allboxes_project.xlsx")
    df.to_excel(excel_file_path, index=False)  
    
    print("Resultados exportados a allboxes_project.xlsx")       


def get_boxesfilteredbyPJ(ruta_archivo):
    
    allboxes_info=boxes_info_joint(ruta_archivo)

    parents_únicos=set()

    allboxes_info_filtered=[]

    for info in allboxes_info:
        parent=info[1]
        #print(parent)
        
        if parent not in parents_únicos:
            parents_únicos.add(parent)
            #print(parents_únicos)
            allboxes_info_filtered.append(info)
                
    return allboxes_info_filtered


def export_excel_infofilteredboxes(rutaifc):
    
    datos_export=get_boxesfilteredbyPJ(rutaifc)
            
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(datos_export, columns=["r_guid", "parent_joint_id", "JointTypeID", "CGT", "QU_Length", "Box_type", "Inst A", "Inst B"])
    # Exportar el DataFrame a un archivo de Excel
    export_path="C:/Users/Adrian Moreno/Desktop"
    excel_file_path=os.path.join(export_path, "filteredboxes_project.xlsx")
    df.to_excel(excel_file_path, index=False)  
    
    print("Resultados exportados a filteredboxes_project.xlsx")

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
            balcony.append([EI_OpeningType,EI_HostComponentInstanceID,QU_Height])
    return balcony
    #for i in balcony[1]: print (i)
    #print (balcony)
    
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
        Params_info=[EI_HostComponentInstanceID,EI_OpeningType, Category]
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
        Params_info=[EI_HostComponentInstanceID,EI_OpeningType, Category]
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
        instancia=balcony[1]
        if instancia in recuento_balconerasporinstancia:
            recuento_balconerasporinstancia[instancia] += 1
        else:
            recuento_balconerasporinstancia[instancia] = 1
            
    for door in alldoors_info:
        instancia=door[0]
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
        PJ=i[6]
        if i[5]=="H.ST_Bottom":
            if PJ in balcony_instances_input.keys():
                balconerasdelainstancia=balcony_instances_input[PJ]
                i.append(balconerasdelainstancia)
                boxesandbalconys.append(i)
                #print(i)
            else:
                i.append('0')
                boxesandbalconys.append(i)
        else:
            i.append('0')
            boxesandbalconys.append(i)
    return boxesandbalconys
            
            
        
def export_excel_infofilteredboxes_withnrofbalconies(rutaifc):
    
    datos_export=getboxesfilteredwithbalconies(rutaifc)
            
    # Crear un DataFrame de pandas con los resultados
    df = pd.DataFrame(datos_export, columns=["r_guid", "parent_joint_id", "JointTypeID", "CGT", "QU_Length", "Box_type", "Inst A", "Inst B", "nr of wind height>1.8 m"])
    # Exportar el DataFrame a un archivo de Excel
    export_path="C:/Users/Adrian Moreno/Desktop"
    excel_file_path=os.path.join(export_path, "filteredboxeswithbalconies_project.xlsx")
    df.to_excel(excel_file_path, index=False)  
    
    print("Resultados exportados a filteredboxeswithbalconies_project.xlsx")            
    
    






"""
def ifc_get_boxes(ifc):
            
    building_elem_proxy_collector = ifc.by_type('IfcBuildingElementProxy')
    joints = []
    joint_info = []
    for i in building_elem_proxy_collector:
        i_psets = ifcopenshell.util.element.get_psets(i)
        if 'EI_Elements Identification' in i_psets:
            eitype = i_psets['EI_Elements Identification'].get('EI_Type', '')
            if 'Joint' in eitype:
                joints.append(i)
    
    for i in joints:
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
        joint_info.append([r_guid, parent_joint_id, JointTypeID, cgt, QU_Length, Box_type])
    
    #return joint_info
    
    pregunta=input("Quieres exportar todas las cajas en un excel?: ")
    if pregunta == "Yes":
        # Crear un DataFrame de pandas con los resultados
        df = pd.DataFrame(joint_info, columns=["r_guid", "parent_joint_id", "JointTypeID", "CGT", "QU_Length", "Box_type"])  

        # Exportar el DataFrame a un archivo de Excel
        #df.to_excel("resultados_ifc.xlsx", index=False)
        export_path=r"C:/Users/Adrian Moreno/Desktop"
        excel_file_path=os.path.join(export_path, "allboxes_project.xlsx")
        df.to_excel(excel_file_path, index=False)

        print("Resultados exportados a allboxes_project.xlsx")
    else:
        return joint_info
"""






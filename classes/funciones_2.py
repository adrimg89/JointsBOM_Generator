import requests
from passwords import *


class Airtable:
    def __init__(self, token, base_id):
        self.token = token
        self.base_id = base_id

    def list(self, table, max_records=None, view='API',fields=None,filter=None,sort=None):
        url = f'https://api.airtable.com/v0/{self.base_id}/{table}'
        params = {
            'maxRecords': max_records,
            'view': view
        }
        if filter is not None:
            params['filterByFormula'] = filter
        if fields is not None:
            params['fields[]'] = fields
        if sort is not None:
            params['sort[]'] = sort
        data = []

        while True:
            response = requests.get(url, params=params, headers={'Authorization': 'Bearer ' + self.token})
            #print (response.url)
            response_data = response.json()
            data.extend(response_data.get('records', []))
            offset = response_data.get('offset', None)
            #print (offset)
            if not offset:
                break
            params['offset'] = offset
            
        return data
        # example with custom filter: at.list('Contracts',filter="AND({Contract Type}='Obra' , {Project}='5')")

    def insert(self, table, data):
        url = f'https://api.airtable.com/v0/{self.base_id}/{table}'
        response = requests.post(url, json={'fields': data}, headers={'Authorization': 'Bearer ' + self.token})
        return response.json()
    
def airtable_conection(api_key,base_id):
    connection=Airtable(api_key,base_id)    
    return connection

jointsplayground_connect = airtable_conection(adri_jointsplayground_api_key,joints_base_id)
materials_connect = airtable_conection(adri_materiales_api_key,materiales_base_id)
joints3playground_connect=airtable_conection(adri_joints3playground_api_key,materiallayers_base_id)

def jointsplayground_connectiongroup():    
    resultados_cgtype = jointsplayground_connect.list(cgtype_table,view='AM', fields=['cgtype_id','Description','api_ConnectionGroup_Class','RL_ConnectionGroupType_ConnectionType_api'])
    connectiongroup_types = []
    for i in resultados_cgtype:
        connectiongroup_types.append(i['fields'])
    return connectiongroup_types

def connectiontype_of_connectiongroup(connectiongroup_type):
    resultados_rl_cgtype_ctype = jointsplayground_connect.list(rl_cgtype_ctype_table, view='AMG_Export', fields=['connectiongroup_type_id','connection_type','Performance','Calculation Formula'],filter="SEARCH('"+connectiongroup_type+"', {connectiongroup_type_id})")
    connection_types = []    
    for i in resultados_rl_cgtype_ctype:
        connection_types.append(i['fields'])
    return connection_types
    #tipo de respuesta: [{'Calculation Formula': 'Fix value', 'Performance': 3, 'connection_type': 'H_T3-0003', 'connectiongroup_type_id': 'CG_0004'}, {'Calculation Formula': 'Length * performance', 'Performance': 5, 'connection_type': 'H_C1-0023', 'connectiongroup_type_id': 'CG_0004'}]

def connectionlayers_of_connectiontype(connection_type):
    resultados_connectionlayers=jointsplayground_connect.list(clayers_table,fields=['connection_type_code','material_id','Performance','Calculation Formula','Current_material_cost','Units'],filter="SEARCH('"+connection_type+"', {connection_type_code})")
    clayers=[]
    for i in resultados_connectionlayers:
        clayers.append(i['fields'])
    return(clayers)
    #tipo de respuesta: [{'Performance': 1, 'Calculation Formula': 'Fix value', 'connection_type_code': 'H_T3-0003', 'material_id': 'MBRA0214'}, {'Performance': 60, 'Calculation Formula': 'Fix value', 'connection_type_code': 'H_T3-0003', 'material_id': 'MFIX0578'}]

def materialcost(sku):
    resultados_materiales = materials_connect.list(materials_table, view='API', fields = ['Material_SKU','Current Material Cost','Units'],filter="SEARCH('"+sku+"', {Material_SKU})")
    materials=[]
    for i in resultados_materiales:
        materials.append(i['fields'])
    return materials
    #tipo de respuesta: [{'Material_SKU': 'MBRA0214', 'Current Material Cost': 7.57}]
    
def jlayers(joint):
    resultados_jlayers=jointsplayground_connect.list(jointlayers_table,fields=['joint_type_code','material_id','Performance','Calculation Formula','Current_material_cost'],filter="SEARCH('"+joint+"',{joint_type_code})")
    joints=[]
    for i in resultados_jlayers:
        joints.append(i['fields'])
    return joints
    #tipo de respuesta: {'Performance': 1.05, 'Calculation Formula': 'Length * performance', 'joint_type_code': 'J_0351', 'material_id': 'MJNT4141', 'Current_material_cost': [5.81]}


def rl_cgtype_ctype():
    resultados_rl_cgtype_ctype = jointsplayground_connect.list(rl_cgtype_ctype_table, view='AMG_Export', fields=['connectiongroup_type_id','connection_type','Performance','Calculation Formula', 'is_modeled'])
    connection_types = []    
    for i in resultados_rl_cgtype_ctype:
        connection_types.append(i['fields'])
    return connection_types
    #tipo de respuesta: [{'Calculation Formula': 'Fix value', 'Performance': 3, 'connection_type': 'H_T3-0003', 'connectiongroup_type_id': 'CG_0004'}, {'Calculation Formula': 'Length * performance', 'Performance': 5, 'connection_type': 'H_C1-0023', 'connectiongroup_type_id': 'CG_0004'}]

def connectionlayers():
    resultados_connectionlayers=jointsplayground_connect.list(clayers_table,fields=['connection_type_code','material_id','Performance','Calculation Formula','Current_material_cost','Units','Description (from Material)','Fase'])
    clayers=[]
    for i in resultados_connectionlayers:
        clayers.append(i['fields'])
    return(clayers)
    #tipo de respuesta: [{'Performance': 1, 'Calculation Formula': 'Fix value', 'connection_type_code': 'H_T3-0003', 'material_id': 'MBRA0214'}, {'Performance': 60, 'Calculation Formula': 'Fix value', 'connection_type_code': 'H_T3-0003', 'material_id': 'MFIX0578'}]

def alljlayers():
    resultados_jlayers=jointsplayground_connect.list(jointlayers_table,fields=['joint_type_code','material_id','Performance','Calculation Formula','Current_material_cost','Long_Name (from Material)','Fase'])
    joints=[]
    for i in resultados_jlayers:
        joints.append(i['fields'])
    return joints    
    # tipo de respuesta: {'Performance': 1.05, 'Calculation Formula': 'Length * performance', 'joint_type_code': 'J_0351', 'material_id': 'MJNT4141', 'Current_material_cost': [5.81]}
    
def allmatlayers():
    resultados_matlayers=joints3playground_connect.list(materiallayers_table,view='API',fields=['api_material_group','api_Material','Performance','calculation_formula','current_material_cost','material_description','Fase'])
    matlayers=[]
    for i in resultados_matlayers:
        matlayers.append(i['fields'])
    return matlayers


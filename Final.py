import pandapower as pp
import pandas as pd
import pandapower.shortcircuit as sc
from pandapower.plotting.plotly.mapbox_plot import set_mapbox_token
# please add your mapbox token
set_mapbox_token()
from pandapower.plotting.plotly import simple_plotly
from pandapower.plotting.plotly import pf_res_plotly

# please write the address of excel input file the name of it in the sent fiels is Project V.2 
fadress='C:/'

net = pp.create_empty_network(f_hz=50.0,sn_mva=1)
nb=0

def Buses(all):
    #Defining Buses
    df= pd.read_excel(fadress, sheet_name="Bus", index_col=0)
    for idx in df.index:
        pp.create_bus(net,vn_kv=df.at[idx,"Voltage"],name=df.at[idx,"Bus name"],geodata=(df.at[idx,"location Y"], df.at[idx,"location X"]))
        global nb
        nb=len(df)
    
    #Defining All the Primary Busses
    if all==True:
        df= pd.read_excel(fadress, sheet_name="Primary Buses", index_col=0)
        for idx in df.index:
            pp.create_bus(net,vn_kv=df.at[idx,"Voltage"],name=df.at[idx,"Bus name"],geodata=(df.at[idx,"location Y"], df.at[idx,"location X"]))
            b=int(nb+idx-1)
            pp.create_ext_grid(net, bus=net.bus.index[b], name="infinite", vm_pu=1.00, va_degree=0.0)
    return(nb)  
def Loads(state):
    #Defining Loads  Now or 20 Years
    df=pd.read_excel(fadress, sheet_name="Load", index_col=0)
    if state=="Now":
        for idx in df.index:
            pp.create_load(net,name=df.at[idx,"Name"],bus=df.at[idx,"Bus"],p_mw=df.at[idx,"Pmax"],q_mvar=df.at[idx,"Qmax"],max_p_mw=df.at[idx,"Pmax"],min_p_mw=df.at[idx,"Pmin"],max_q_mvar=df.at[idx,"Qmax"],min_q_mvar=df.at[idx,"Qmin"])
    # withh p=pmax and q=qmax  pp.create_load(net,name=df.at[idx,"Name"],bus=df.at[idx,"Bus"],p_mw=df.at[idx,"Pmax"],q_mvar=df.at[idx,"Qmax"],max_p_mw=df.at[idx,"Pmax"],min_p_mw=df.at[idx,"Pmin"],max_q_mvar=df.at[idx,"Qmax"],min_q_mvar=df.at[idx,"Qmin"])
    elif state=="20years":
        for idx in df.index:
            pp.create_load(net,name=df.at[idx,"Name"],bus=df.at[idx,"Bus"],p_mw=df.at[idx,"Pmax20"],q_mvar=df.at[idx,"Qmax20"],max_p_mw=df.at[idx,"Pmax20"],min_p_mw=df.at[idx,"Pmin20"],max_q_mvar=df.at[idx,"Qmax20"],min_q_mvar=df.at[idx,"Qmin20"])
    return
def Line():
#Defining Lines
    df=pd.read_excel(fadress, sheet_name="Line", index_col=0)
    for idx in df.index:
        pp.create_line_from_parameters(net,name=df.at[idx,"name"],from_bus=df.at[idx,"from_bus"],to_bus=df.at[idx,"to_bus"],length_km=df.at[idx,"length_km"],r_ohm_per_km=df.at[idx,"r_ohm_per_km"],x_ohm_per_km=df.at[idx,"x_ohm_per_km"],c_nf_per_km=df.at[idx,"c_nf_per_km"],r0_ohm_per_km=df.at[idx,"r0_ohm_per_km"],x0_ohm_per_km=df.at[idx,"x0_ohm_per_km"],c0_nf_per_km=df.at[idx,"c0_nf_per_km"],max_i_ka=df.at[idx,"max_i_ka"])
    #print(*df.loc[idx, :])
    return
def rest():
    #Defining External Grids
    pp.create_ext_grid(net, bus=net.bus.index[1], name="P05", vm_pu=1.00, va_degree=0.0,s_sc_max_mva=2400,rx_max=0)
    pp.create_ext_grid(net, bus=net.bus.index[0], name="P04", vm_pu=1.00, va_degree=0.0,s_sc_max_mva=2400,rx_max=0) 
    """ pp.create_ext_grid(net, bus=net.bus.index[nb-1], name="P05", vm_pu=1.00, va_degree=0.0,s_sc_max_mva=2400,rx_max=0)
    pp.create_ext_grid(net, bus=net.bus.index[nb-2], name="P04", vm_pu=1.00, va_degree=0.0,s_sc_max_mva=2400,rx_max=0) """
#Defining Transformers
    pp.create_transformer(net, net.bus.index[nb-1], net.bus.index[1] , std_type="40 MVA 110/20 kV",parallel=2)
    pp.create_transformer(net, net.bus.index[nb-2], net.bus.index[0] , std_type="40 MVA 110/20 kV",parallel=2)
    """     net.trafo.uk_percent[0]=0.1
    net.trafo.uk_percent[1]=0.1 """
    net.trafo.vk_percent[0]=10
    net.trafo.vk_percent[1]=10
    net.trafo.vkr_percent[0]=0.0
    net.trafo.vkr_percent[1]=0.0
    net.trafo.tap_pos[0]=0
    net.trafo.tap_pos[1]=0
    print(net.trafo)
    #print(net.bus)
    return



def Operation_Topology (state):
    if state=='Normal':
        swo = pp.create_switch(net, bus=net.bus.index[9], element=net.line.index[8], et="l", type="LBS", closed=False)
        print(state)
        return()
    elif state== 'Contingency_P04':
        sw4 = pp.create_switch(net, bus=net.bus.index[0], element=net.line.index[16], et="l", type="LBS", closed=False)
        print(state)
        return()
    elif state== 'Contingency_P05':
        sw4 = pp.create_switch(net, bus=net.bus.index[1], element=net.line.index[0], et="l", type="LBS", closed=False)
        print(state)
        return()
def SC_calculation():
    sc.calc_sc(net, case="max",branch_results=True)
    net.res_bus_sc
    net.res_line_sc
def SC_calculation_presentation(line,bus):
    if line==True:
        print(net.res_line_sc)
    if bus==True:
        print(net.res_bus_sc)
    return
def Powerflow_calculation():
    pp.runpp(net,"nr")
def Powerflow_presentation(table,map):
    if table==True:
        print(net.res_line)
    if map==True:
        pf_res_plotly(net,on_map=True,map_style="streets", projection="epsg:3035")
    return
def Operation_Topology_presentation (style):
    simple_plotly(net, on_map=True, map_style=style, projection="epsg:3035")


Buses(all=False)
Loads(state="Now") 
#Loads(state="20years")
Line()
rest()
Operation_Topology(state="Normal")
#Operation_Topology(state='Contingency_P04')
#Operation_Topology(state='Contingency_P05')

#Operation_Topology_presentation(style='streets')
# style can be streets, dark, satellite
SC_calculation()
Powerflow_calculation()
Powerflow_presentation(table=True,map=True)
SC_calculation_presentation(line=True,bus=False)


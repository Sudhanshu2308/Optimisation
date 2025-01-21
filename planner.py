import gurobipy as gp
from gurobipy import *
import random as rd
import numpy as np
import pandas as pd
import folium
import os

state_data = None
time_data = None
xls = {}
total_place = {}
dm_dict = {}
arr = {}
cord = {}
Lat = {}
Long = {}
start_time = {}
end_time = {}

def file_location(folder_path):
    file_locations = {}
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith(".xlsm"):
            base_filename, _ = os.path.splitext(filename)  
            file_path = os.path.join(folder_path, filename)
            file_locations[base_filename] = file_path
    return file_locations

def planning(state,visiting_major_station, stay_duration, Q):
    global state_data,time_data,xls,total_place, dm_dict, arr, cord, Lat, Long
    folder_dist = "./Statewise"
    folder_time = "./Time Window"
    if state_data is None:
        state_data = file_location(folder_dist)
        time_data = file_location(folder_time)
    if state not in xls:
        xls[state] = pd.ExcelFile(state_data[state])
        total_place[state] = {}
        data = {}
        for sheet_name in xls[state].sheet_names:
            data[sheet_name.strip()] = pd.read_excel(state_data[state], sheet_name=sheet_name)
            total_place[state][sheet_name.strip()] = data[sheet_name.strip()].shape[1]-2
        dm_dict[state] = {}
        for sheet_name, df in data.items():
            # Exclude first row and first column
            array = df.iloc[0:, 1:].values
            dm_dict[state][sheet_name] = array.tolist()
        arr[state] = {}
        for sheet_name, df in data.items():
            # Exclude first row and first column
            arr1 = df.iloc[:, 0].values
            arr[state][sheet_name] = arr1.tolist()
        df=pd.read_excel("./Coordinates/State_Tourist_Places_LatLong.xlsx",sheet_name=state)
        temp = df.Places.values.tolist()
        for i in range(len(temp)):
            temp[i] = temp[i].strip()
        cord[state]=temp
        Lat[state]=df.Latitude.values.tolist()
        Long[state]=df.Longitude.values.tolist()
    map = folium.Map(location=[np.mean(Lat[state]), np.mean(Long[state])], 
                    zoom_start=6, 
                    control_scale=True)
    if state not in start_time:
        file =time_data[state]
        files = pd.ExcelFile(file)
        start_time[state] = {}
        end_time[state] = {}
        for sheet in files.sheet_names:
            df = pd.read_excel(file, sheet_name = sheet)
            start_time[state][sheet]=df["start time"].tolist()
            end_time[state][sheet]=df["end time"].tolist()
    M = 10000
    out = []
    for ia in visiting_major_station:
        p = total_place[state][ia]
        d = stay_duration[ia]
        open_time = start_time[state][ia]
        close_time = end_time[state][ia]
        mat = dm_dict[state][ia]
        P = [i for i in range(p)] 
        T = [i for i in range(1,d+1)] 
        tm = [[round(elem/50,2) for elem in row] for row in mat]
        
        m=gp.Model('route')
        
        #Variable
        xijt = m.addVars(P, P, T, vtype=GRB.BINARY, name ='xijt')
        xt = m.addVars(T, vtype=GRB.BINARY, name='xt')
        sit = m.addVars(P,T,vtype = GRB.CONTINUOUS, name = 'sit')
        
        #Objective Function
        obj_fn = sum(xijt[i,j,t] * mat[i][j] for t in T for i in P for j in P)
        m.setObjective(obj_fn, GRB.MINIMIZE)
        
        m.addConstrs(gp.quicksum(xijt[i,i,t] for t in T) == 0 for i in P)
        m.addConstrs(gp.quicksum(xijt[i,j,t] for i in P) == gp.quicksum(xijt[j,i,t] for i in P) for j in P for t in T)
        m.addConstrs(gp.quicksum(xijt[i,j,t] for i in P for t in T) == 1 for j in range(1,p))
        m.addConstrs(gp.quicksum(xijt[0,j,t] for j in range(1,p)) == xt[t] for t in T)
        m.addConstrs(xt[t] >= xt[t+1] for t in range(1,len(T)))
        m.addConstrs(gp.quicksum(xijt[i,j,t]*tm[i][j] for i in P for j in range(1,p) ) <= Q for t in T)

        #Miller-Tucker-Zemlin formulation Subtour Elimination Constraints
        q = [rd.randint(3, 3) for i in range(p)]  # generate the random vector
        
        #Time-Window Constraints
        m.addConstrs(sit[i,t] + q[i] + tm[i][j] - sit[j,t] <= M*(1-xijt[i,j,t]) for i in range(1,len(P)) for j in range(1,len(P)) for t in T if i!=j)
        for i in P:  # Iterate over all i
            for t in T:  # Iterate over all k
                m.addConstr(sit[i, t] <= (close_time[i]-q[i]) * quicksum(xijt[i, j, t] for j in P))
        for i in P:  # Iterate over all i
            for t in T:  # Iterate over all k
                m.addConstr(sit[i, t] >= open_time[i] * quicksum(xijt[i, j, t] for j in P))
        u = m.addVars(P, name ="u")
        u[0].LB = 0
        u[0].UB = 0
        for i in P :
            u[i].LB = min([(q[i]+tm[j][i]) for j in P for t in T])
            u[i].UB = Q
        m.addConstrs( u[i] - u[j] + Q * xijt[i,j,t] <= Q - q[j] for i in P for j in P for t in T if j != 0 )
        
        m.update()pip
        m.optimize()
        A = [(i, j, t) for i in P for j in P for t in T if i != j] # matrix distance
        mi = []
        if m.status == GRB.OPTIMAL:
            dict = {}
            for t in T:
                for j in P:
                    for i in P:
                        if xijt[i, j, t].x > 0.99:
                            if t not in dict:
                                dict[t] = {}
                            dict[t][i] = j
            for t in T:
                start = 0
                curr = 0
                linear_graph = []
                while dict[t][curr] != start:
                    linear_graph.append(arr[state][ia][curr])
                    curr = dict[t][curr]
                linear_graph.append(arr[state][ia][curr])
                linear_graph.append(arr[state][ia][start])
                dict[t] = linear_graph
            out.append(dict)
            for a in T:
                aa = [(i, j) for i in P for j in P if xijt[i, j, a].x > 0.99]
                mi.append(aa)
        dict1 = {}
        for idx, arc in enumerate(mi):        
            dict1[idx+1] = arc
        place = arr[state][ia]
        col = ['red', 'green', 'darkblue', 'purple','red', 'green']
        route = []
        for key in dict1:
            for tuple1 in dict1[key]:
                x,y =  tuple1
                if x == 0:
                    r = arr[state][ia][x]
                    folium.Circle(location=([float(Lat[state][cord[state].index(r)]),float(Long[state][cord[state].index(r)])]),radius=80,color="crimson",fill=False).add_to(map)
                if y == 0:
                    r = arr[state][ia][y]
                    folium.Circle(location=([float(Lat[state][cord[state].index(r)]),float(Long[state][cord[state].index(r)])]),radius=80,color="crimson",fill=False).add_to(map) 
                else:
                    i = arr[state][ia][x]
                    j = arr[state][ia][y]                
                    folium.Circle(location=([float(Lat[state][cord[state].index(i)]),float(Long[state][cord[state].index(i)])]),radius=40,color="black",fill=False).add_to(map)
                    folium.Circle(location=([float(Lat[state][cord[state].index(j)]),float(Long[state][cord[state].index(j)])]),radius=40,color="black",fill=False).add_to(map)
                if arr[state][ia][x] in cord[state] and arr[state][ia][y] in cord[state]:
                    i = arr[state][ia][x]
                    j = arr[state][ia][y]
                    folium.PolyLine(locations=[[float(Lat[state][cord[state].index(i)]),float(Long[state][cord[state].index(i)])],[float(Lat[state][cord[state].index(j)]),float(Long[state][cord[state].index(j)])]]
                                ,color=col[key], weight=1.5).add_to(map)                
    # in this there are mainly 2 problems 
    # 1. u[i] should have time taken to reach i from any j, but instead it is only a 1d list generated randomly.
    # 2. if a station has only one place and 2 days are allotted error is coming as we re saying that it should use all days
    # 3. prioritizing days is not being done. for example: if there are 7 palaces,
    #     on day 1 : 0-5-0
    #     on day 2 : 0-4-6-0
    #     on day 3 : 0-1-3-2-0    
    map.save('map.html')
    return out


    # if xls is None:
    #     xls = pd.ExcelFile(excel_file)
    #     stnames = []
    #     pl = []
    #     data = {}
    #     for sheet_name in xls.sheet_names:
    #         data[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)
    #         stnames.append(sheet_name)
    #     dm = []
    #     dist=[]
    #     places=[]
    #     arrays = {}
    #     for sheet_name, df in data.items():
    #         # Exclude first row and first column
    #         array = df.iloc[0:,1:].values
    #         arrays[sheet_name] = array
    #     for sheet_name in stnames:
    #         array = arrays[sheet_name]
    #         dist = array.tolist()
    #         pl.append(len(dist)-1)
    #         dm.append(dist)
        
    #     arr = {}
    #     for sheet_name, df in data.items():
    #         # Exclude first row and first column
    #         arr1 = df.iloc[:, 0].values
    #         arr[sheet_name] = arr1

    #     for sheet_name in stnames:
    #         arr1 = arr[sheet_name]
    #         pla = arr1.tolist()
    #         places.append(pla)

    #     df=pd.read_excel('Coordinates.xlsx',sheet_name='allcoord')
    #     cord=df.id.values.tolist()
    #     Lat=df.xcoord.values.tolist()
    #     Long=df.ycoord.values.tolist()

    # # shuffle the stations days as per stnames
    # new_days = [0]*len(stnames)
    # for i in range(len(days)):
    #     new_days[stnames.index(station_names[i])] = days[i]
    # print(places)
    # df.columns = ['id','Lat','Long']
    # map = folium.Map(location=[df.Lat.mean(), df.Long.mean()], 
    #             zoom_start=1, 
    #             control_scale=True)
    # out = []
    # for ia in range(len(stnames)):
    #     p = pl[ia]
    #     d = days[ia]
    #     mat = dm[ia]
    #     P = [i for i in range(p)]
    #     T = [i for i in range(1,d+1)]
    #     tm = [[round(elem/50,2) for elem in row] for row in mat]

    # #     tm = [[round(elem/50 +2,2) for elem in row] for row in mat]
    # #     for i in range(len(tm)):
    # #         tm[i][i] -= 2
    # #     print(tm)
    #     print("\n")

    #     m=gp.Model('route')

    #     #Variable
    #     xijt = m.addVars(P, P, T, vtype=GRB.BINARY, name ='xijt')

    #     #Objective Function
    #     obj_fn = sum(xijt[i,j,t] * mat[i][j] for t in T for i in P for j in P)
    #     m.setObjective(obj_fn, GRB.MINIMIZE)

    # #     #Objective Function changed
    # #     obj_fn = sum(xijt[i,j,t] for t in T for i in P for j in P)
    # #     m.setObjective(obj_fn, GRB.MAXIMIZE)


    #     m.addConstrs(gp.quicksum(xijt[i,i,t] for t in T) == 0 for i in P)
    #     m.addConstrs(gp.quicksum(xijt[i,j,t] for i in P) == gp.quicksum(xijt[j,i,t] for i in P) for j in P for t in T)
    #     m.addConstrs(gp.quicksum(xijt[i,j,t] for i in P for t in T) == 1 for j in range(1,p))
    #     m.addConstrs(gp.quicksum(xijt[0,j,t] for j in range(1,p)) == 1 for t in T)
    # #order check for i,j
    #     m.addConstrs(gp.quicksum(xijt[i,j,t]*tm[i][j] for i in P for j in range(1,p) ) <= 15 for t in T)


    # #     new constraints


    # #Miller-Tucker-Zemlin formulation Subtour Elimination Constraints
    #     q = [rd.randint(3, 3) for i in range(p)]  # generate the random vector
    #     u = m.addVars(P)
    #     u[0].LB = 0
    #     u[0].UB = 0
    #     for i in P :

    # #             u[i].LB = q[i]
    #         u[i].LB = min([(q[i]+tm[j][i]) for j in P for t in T])
    #         u[i].UB = Q
    # #     print([(q[i]+tm[j][i])* xijt[i,j,t] for i in P for j in P for t in T])
    #     m.addConstrs( u[i] - u[j] + Q * xijt[i,j,t] <= Q - q[j] for i in P for j in P for t in T if j != 0 )



    # # #     m.remove(c)  # remove the previous MTZ constraints
    # #     q[0] = 0
    # #     m.addConstrs( u[i] - u[j] + Q * xijt[i,j,t] + ( Q - q[i] - q[j] ) * xijt[j,i,t] <= Q - q[j] for i in P for j in P for t in T if j != 0 )


    #     m.optimize()
    #     A = [(i, j, t) for i in P for j in P for t in T if i != j] # matrix distance
    #     mi = []

    #     if m.status == GRB.OPTIMAL:
    #         # Write the m.X values in Dict
    #         dict = {}
    #         for t in T:
    #             for j in P:
    #                 for i in P:
    #                     if xijt[i, j, t].x > 0.99:
    #                         if t not in dict:
    #                             dict[t] = {}
    #                         dict[t][i] = j
    #         # print(dict)
    #         m.printAttr('X')
    #         for t in T:
    #             start = 0
    #             curr = 0
    #             linear_graph = []
    #             while dict[t][curr] != start:
    #                 # print(curr, end=" ")
    #                 linear_graph.append(places[ia][curr])
    #                 curr = dict[t][curr]
    #             # print()
    #             linear_graph.append(places[ia][curr])
    #             linear_graph.append(places[ia][start])
    #             dict[t] = linear_graph
    #         out.append(dict)
    #         for a in T:
    #             aa = [(i, j) for i in P for j in P if xijt[i, j, a].x > 0.99]
    #             mi.append(aa)
    #     else:
    #         print("Model did not solve successfully.")
    #     dict1 = {}
    #     for idx, arc in enumerate(mi):
    #         dict1[idx+1] = arc

    #     # dict1
    #     # print("s")
    #     # print(s)

    # #     map_stats = folium.Map()
    #     col = ['red', 'green', 'darkblue', 'purple','red', 'green']

    #     for key in dict1:
    #         for tuple1 in dict1[key]:
    #             x,y =  tuple1
    #             if x == 0:
    #                 r = places[ia][x]
    #                 folium.Circle(location=([float(Lat[cord.index(r)]),float(Long[cord.index(r)])]),radius=80,color="crimson",fill=False).add_to(mapap)
    #             if y == 0:
    #                 r = places[ia][y]
    #                 folium.Circle(location=([float(Lat[cord.index(r)]),float(Long[cord.index(r)])]),radius=80,color="crimson",fill=False).add_to(mapap)

    #             else:
    #                 i = places[ia][x]
    #                 j = places[ia][y]
    #                 folium.Circle(location=([float(Lat[cord.index(i)]),float(Long[cord.index(i)])]),radius=40,color="black",fill=False).add_to(mapap)
    #                 folium.Circle(location=([float(Lat[cord.index(j)]),float(Long[cord.index(j)])]),radius=40,color="black",fill=False).add_to(mapap)


    #             # print("x =", x, "y =", y)
    #             if places[ia][x] in cord and places[ia][y] in cord:
    #                 i = places[ia][x]
    #                 j = places[ia][y]
    #                 # print(i)
    #                 # print(j)
    # #                 if i==j:
    # #                     continue
    # #                 else:

    #                 folium.PolyLine(locations=[[float(Lat[cord.index(i)]),float(Long[cord.index(i)])],[float(Lat[cord.index(j)]),float(Long[cord.index(j)])]]
    #                             ,color=col[key], weight=1.5).add_to(mapap)

    # mapap.save('map.html')
    # return stnames, new_days, mapap._repr_html_(), out

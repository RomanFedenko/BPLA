from openpyxl.chart import BarChart,Reference,PieChart
import matplotlib.pyplot as plt

class Bpla:
    def __init__(self,uniq_id,name,model,mac_number,max_speed,max_flight_time,max_flight_dist,weight):
        self.uniq_id = uniq_id
        self.name = name
        self.model = model
        self.mac_number = mac_number
        self.max_speed=max_speed
        self.max_flight_time=max_flight_time
        self.max_flight_dist=max_flight_dist
        self.weight = weight

    def add_data(self,data_bd,connect,cursor):
        if str(data_bd[0]) and str(data_bd[1]) and int(data_bd[2]) and float(data_bd[3]) and float(data_bd[4]) and float(data_bd[5]) and float(data_bd[6]):
                self.name = data_bd[0]
                self.model = data_bd[1]
                self.mac_number = data_bd[2]
                self.max_speed = data_bd[3]
                self.max_flight_time = data_bd[4]
                self.max_flight_dist = data_bd[5]
                self.weight = data_bd[6]
                cursor.execute('''INSERT INTO public.bpla (name, model, mac_number, max_speed, max_flight_time, max_flight_dist, weight)
                                  VALUES (%s,%s,%s,%s,%s,%s,%s)''', (self.name, self.model, self.mac_number, self.max_speed, self.max_flight_time, self.max_flight_dist, self.weight))
                connect.commit()
                print("Data added successfully")
        else:
            print("Incorrect input data")


    def list(self,cursor):
        cursor.execute('''SELECT * FROM public.bpla''')
        data = cursor.fetchall()
        return data




    def delete(self,identificator,connect,cursor):
        cursor.execute('''SELECT * FROM public.bpla WHERE uniq_id=%s''',(identificator,))
        row = cursor.fetchone()
        if row is not None:
            cursor.execute('''DELETE FROM public.bpla WHERE uniq_id=%s''',(identificator,))
            connect.commit()
            print("Delete is successfully")
        else:
            print("No such id found")

    def edit(self,mas,index,connect,cursor):
        if type(mas[0])==str and type(mas[1])==str and type(mas[2])==int and type(mas[3])==float and type(mas[4])==float and type(mas[5])==float and type(mas[6])==float:
            cursor.execute('''SELECT * FROM public.bpla WHERE uniq_id=%s''',(index,))
            row = cursor.fetchone()
            if row is not None:
                cursor.execute('''UPDATE public.bpla SET name=%s,model=%s,mac_number=%s,max_speed=%s,max_flight_time=%s,max_flight_dist=%s,weight=%s WHERE uniq_id=%s''',(mas[0],mas[1],mas[2],mas[3],mas[4],mas[5],mas[6],index))
                connect.commit()
                print("Edit complete")
            else:
                print("No such id found")
        else:
            print("Incorrect input data")

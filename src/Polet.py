
class Polet:
    def __init__(self,uniq_id,uniq_id_bpla,flight_time,flight_dist,flight_height,task_kol_complete,bpla_notValid,fuel_for_flight):
        self.uniq_id = uniq_id
        self.uniq_id_bpla = uniq_id_bpla
        self.flight_time = flight_time
        self.flight_dist = flight_dist
        self.flight_height = flight_height
        self.task_kol_complete = task_kol_complete
        self.bpla_notValid = bpla_notValid
        self.fuel_for_flight = fuel_for_flight


    def add(self,mas,sost_bpla,cursor,connect):
        if (int(mas[0]) and float(mas[1]) and float(mas[2]) and float(mas[3])
                and int(mas[4]) and bool(sost_bpla) and float(mas[5])):
            self.uniq_id_bpla = mas[0]
            self.flight_time = mas[1]
            self.flight_dist = mas[2]
            self.flight_height = mas[3]
            self.task_kol_complete = mas[4]
            self.bpla_notValid = sost_bpla
            self.fuel_for_flight = mas[5]
            cursor.execute('''INSERT INTO public.polet (uniq_id_bpla,flight_time,flight_dist,flight_height,task_kol_complete,faulty_bpla,fuel_spent_on_flight)
             VALUES (%s,%s,%s,%s,%s,%s,%s)''',(self.uniq_id_bpla,self.flight_time,self.flight_dist,self.flight_height,self.task_kol_complete,self.bpla_notValid,self.fuel_for_flight))
            connect.commit()
            print("Polet added successfully")
        else:
            print("Incorrect input data")


    def list(self, cursor):
        cursor.execute('''SELECT * FROM public.polet''')
        data = cursor.fetchall()
        return data

    def delete(self,data_bd,cursor,connect):
        cursor.execute('''SELECT * FROM public.polet WHERE uniq_id_bpla=%s''',(data_bd,))
        row = cursor.fetchone()
        if row is not None:
            cursor.execute('''DELETE FROM public.polet WHERE uniq_id_bpla=%s''',(data_bd,))
            connect.commit()
            print("Polet delete successfully")
        else:
            print("Noc suh id found")


    def edit(self,data_bd,cursor,connect):
        if type(data_bd[0])==int and type(data_bd[1]) == int and type(data_bd[2]) == float and type(data_bd[3]) == float and type(data_bd[4]) == float and type(
                data_bd[5]) == int and type(data_bd[6]) == bool and type(data_bd[7]) == float:
            cursor.execute('''SELECT * FROM public.polet WHERE uniq_id=%s''',(data_bd[0],))
            row = cursor.fetchone()
            if row is not None:
                cursor.execute('''UPDATE public.polet SET flight_dist=%s WHERE uniq_id=%s AND uniq_id_bpla=%s AND flight_time=%s
                            AND flight_dist=%s AND flight_height=%s AND task_kol_complete=%s AND faulty_bpla=%s AND fuel_spent_on_flight=%s''',
                            (5000,data_bd[0],data_bd[1],data_bd[2],data_bd[3],data_bd[4],data_bd[5],data_bd[6],data_bd[7]))
                connect.commit()
                print("Polet edit successfully")
            else:
                print("No such id found")
        else:
            print("Incorrect input number")

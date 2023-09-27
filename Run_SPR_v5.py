import numpy as np
import pandas as pd
import os
import matplotlib.pyplot as plt
import dearpygui.dearpygui as dpg
from newportxps import NewportXPS
import time

from ctypes import *
import sys
import time

class CommandError(Exception):
    '''The function in the usbdll.dll was not sucessfully evaluated'''

class Newport_XPS:
    def __init__(self, group):
        self.xps = NewportXPS('192.168.254.254', username='Administrator', password='Administrator',timeout=1)
        print(self.xps.status_report())

        for gname, info in self.xps.groups.items():
            print(gname, info)

        for sname, info in self.xps.stages.items():
            print(sname, self.xps.get_stage_position(sname), info)

        self.xps.kill_group(group)
        self.xps.initialize_group(group)

    def home_group(self, group):
        self.xps.home_group(group)

    def move_abs(self, group_name, pos):
        self.xps.move_group(group=group_name, **{"x":pos,"y":pos})

        return

class Newport_2936:
    def __init__(self,interval_ms=1,buff_size=1000,wavelength=633):
        try:
            self.LIBNAME = r"C:\Program Files\Newport\Newport USB Driver\Bin\usbdll.dll"
            self.lib = windll.LoadLibrary(self.LIBNAME)
            self.product_id = 0xCEC7
        except WindowsError as e:
            print(e.strerror)
            sys.exit(1)

        self.open_device_all_products_all_devices()
        self.open_device_with_product_id()
        # here instrument[0] is the device id, [1] is the model number and [2] is the serial number
        self.instrument = self.get_instrument_list()
        [self.device_id, self.model_number, self.serial_number] = self.instrument

        self.wavelength=wavelength
        self.set_wavelength(self.wavelength)
        self.interval_ms=interval_ms
        self.buff_size=buff_size

        reading_freq=int(self.interval_ms * 10)


        self.write('PM:DS:BUF ' + str(1))  # set to act as ring buffer, where oldest values are overwritten  once buffer fills. Means average is always average of most recent period of time.
        self.write('PM:DS:INT ' + str(reading_freq))  # measurement mode of 'CW continuous' puts measurements in the buffer at rate of 0.1ms, so here we take every nth reading
        self.write('PM:DS:SIZE ' + str(self.buff_size))
        self.write('PM:DS:EN 1') # triggers start of data collection

    def open_device_all_products_all_devices(self):
        """
        
        """
        status = self.lib.newp_usb_init_system()  # SHould return a=0 if a device is connected
        if status != 0:
            raise CommandError()
        else:
            print('Success!! your are conneceted to one or more of Newport products')

    def open_device_with_product_id(self):
        """
        opens a device with a certain product id
        """
        cproductid = c_int(self.product_id)
        useusbaddress = c_bool(1)  # We will only use deviceids or addresses
        num_devices = c_int()
        try:
            status = self.lib.newp_usb_open_devices(
                cproductid, useusbaddress, byref(num_devices))

            if status != 0:
                self.status = 'Not Connected'
                raise CommandError(
                    "Make sure the device is properly connected")
            else:
                print('Number of devices connected: ' + str(num_devices.value) + ' device/devices')
                self.status = 'Connected'
        except CommandError as e:
            print(e)
            sys.exit(1)

    def close_device(self):
        """
        Closes the device
        :raise CommandError:
        """
        status = self.lib.newp_usb_uninit_system()  # closes the units
        if status != 0:
            raise CommandError()
        else:
            print('Closed the newport device connection. Have a nice day!')

    def get_instrument_list(self):
        arInstruments = c_int()
        arInstrumentsModel = c_int()
        arInstrumentsSN = c_int()
        nArraySize = c_int()
        try:
            status = self.lib.GetInstrumentList(byref(arInstruments), byref(arInstrumentsModel), byref(arInstrumentsSN),
                                                byref(nArraySize))
            if status != 0:
                raise CommandError('Cannot get the instrument_list')
            else:
                instrument_list = [arInstruments.value,
                                   arInstrumentsModel.value, arInstrumentsSN.value]
                print('Arrays of Device Id\'s: Model number\'s: Serial Number\'s: ' + str(instrument_list))
                return instrument_list
        except CommandError as e:
            print(e)

    def ask(self, query_string):
        """
        Write a query and read the response from the device
        :rtype : String
        :param query_string: Check Manual for commands, ex '*IDN?'
        :return: :raise CommandError:
        """
        status = -1
        query = create_string_buffer(query_string.encode('ascii'))
        leng = c_ulong(sizeof(query))
        cdevice_id = c_long(self.device_id)
        status = self.lib.newp_usb_send_ascii(
            self.device_id, byref(query), leng)
        if status != 0:
            raise CommandError(
                'Something appears to be wrong with your query string')
        else:
            pass

        time.sleep(0.1)
        response = create_string_buffer(1024)
        leng = c_ulong(1024)
        read_bytes = c_ulong()
        status = self.lib.newp_usb_get_ascii(
            cdevice_id, byref(response), leng, byref(read_bytes))
        
        if status != 0:
            raise CommandError(
                'Connection error or Something apperars to be wrong with your query string')
        else:
            answer = response.value[0:read_bytes.value].decode('ascii')
            answer = answer.rstrip('\r\n')
        return answer

    def write(self, command_string):
        """
        Write a string to the device
        :param command_string: Name of the string to be sent. Check Manual for commands
        :raise CommandError:
        """
        command = create_string_buffer(bytes(command_string, 'ascii'))
        length = c_ulong(sizeof(command))
        cdevice_id = c_long(self.device_id)
        status = self.lib.newp_usb_send_ascii(
            cdevice_id, byref(command), length)
        try:
            if status != 0:
                raise CommandError(
                    'Connection error or  Something apperars to be wrong with your command string')
            else:
                pass
        except CommandError as e:
            print(e)

    def set_wavelength(self, wavelength):
        """
        Sets the wavelength on the device
        :param wavelength: float
        """
        if isinstance(wavelength, float) == True:
            print('Warning: Wavelength has to be an integer. Converting to integer')
            wavelength = int(wavelength)

        if wavelength >= int(self.ask('PM:MIN:Lambda?')) and wavelength <= int(self.ask('PM:MAX:Lambda?')):
            self.write('PM:Lambda ' + str(wavelength))
        else:
            print('Wavelenth out of range, use the current lambda')

    def set_filtering(self, filter_type=0):
        """
        Set the filtering on the device
        :param filter_type:
        0:No filtering
        1:Analog filter
        2:Digital filter
        3:Analog and Digital filter
        """
        if isinstance(filter_type, int) == True:
            if filter_type == 0:
                self.write('PM:FILT 0')  # no filtering
            elif filter_type == 1:
                self.write('PM:FILT 1')  # Analog filtering
            elif filter_type == 2:
                self.write('PM:FILT 2')  # Digital filtering
            elif filter_type == 1:
                self.write('PM:FILT 3')  # Analog and Digital filtering

        else:  # if the user gives a float or string
            print('Wrong datatype for the filter_type. No filtering being performed')
            self.write('PM:FILT 0')  # no filtering

    def read_buffer(self):
        """
        Stores the mean and standard deviation power values at a certain wavelength.
        :param wavelength: float: Wavelength at which this operation should be done. float.
        :param buff_size: int: nuber of readings that will be taken
        :param interval_ms: float: Time between readings in ms.
        :return: [actualwavelength,mean_power,std_power]
        """
        
        # self.write('PM:DS:EN 1') # triggers start of data collection
        # Waits for the buffer is full or not.
        # while int(self.ask('PM:DS:COUNT?')) < number_values:
        #     time.sleep(0.001 * self.interval_ms * number_values / 10) # Not needed anymore as we are using a ring buffer.
        # actualwavelength = self.ask('PM:Lambda?')
        answer=self.ask('PM:STAT:MEAN?;PM:STAT:SDEV?')
        answer=[float(i) for i in answer.split(',')]
        # std_power = self.ask('PM:STAT:SDEV?')
        # self.write('PM:DS:Clear')
        return answer

    def read_instant_power(self, wavelength=700):
        """
        reads the instanenous power
        :param wavelength:
        :return:[actualwavelength,power]
        """
        self.set_wavelength(wavelength)
        actualwavelength = self.ask('PM:Lambda?')
        power = self.ask('PM:Power?')
        return [actualwavelength, power]

    def sweep(self, swave, ewave, interval, buff_size=1000, interval_ms=1):
        """
        Sweeps over wavelength and records the power readings. At each wavelength many readings can be made
        :param swave: int: Start wavelength
        :param ewave: int: End Wavelength
        :param interval: int: interval between wavelength
        :param buff_size: int: nunber of readings
        :param interval_ms: int: Time betweem readings in ms
        :return:[wave,power_mean,power_std]
        """
        self.set_filtering()  # make sure their is no filtering
        data = []
        num_of_points = (ewave - swave) / (1 * interval) + 1

        for i in np.linspace(swave, ewave, int(num_of_points)).astype('int'):
            data.extend(self.read_buffer())
        data = [float(x) for x in data]
        wave = data[0::3]
        power_mean = data[1::3]
        power_std = data[2::3]
        return [wave, power_mean, power_std]

    def sweep_instant_power(self, swave, ewave, interval):
        """
        Sweeps over wavelength and records the power readings. only one reading is made
        :param swave: int: Start wavelength
        :param ewave: int: End Wavelength
        :param interval: int: interval between wavelength
        :return:[wave,power]
        :return:
        """
        self.set_filtering(self.device_id)  # make sure there is no filtering
        data = []
        num_of_points = (ewave - swave) / (1 * interval) + 1
        import numpy as np

        for i in np.linspace(swave, ewave, int(num_of_points)).astype(int):
            data.extend(self.read_instant_power(i))
        data = [float(x) for x in data]
        wave = data[0::2]
        power = data[1::2]
        return [wave, power]

    def plotter_instantpower(self, data):
        plt.close('All')
        plt.plot(data[0], data[1], '-ro')
        plt.show()

    def plotter(self, data):
        plt.close('All')
        plt.errorbar(data[0], data[1], data[2], fmt='ro')
        plt.show()

    def plotter_spectra(self, dark_data, light_data):
        plt.close('All')
        plt.errorbar(dark_data[0], dark_data[1], dark_data[2], fmt='ro')
        plt.errorbar(light_data[0], light_data[1], light_data[2], fmt='go')
        plt.show()

class QueueElement:
    def __init__(self, name, func, duration):
        self.name=name
        self.func=func
        self.duration=duration
        
    def execute(self):
        self.func()
        
# ===================================

def xps_home_callback(sender, data):

    pos = dpg.get_value(dpg_move_pos)
    xps1.home_group("XY")
    dpg.configure_item(dpg_move, enabled=True)
    dpg.configure_item(dpg_run, enabled=True)
    dpg.set_value(dpg_move_pos, 90)
    rel_pos[0] = 0
    rel_pos_str = f"Position: [{str(rel_pos[0])}]"
    dpg.set_value(dpg_rel_pos_str, rel_pos_str)

def xps_move_callback(sender, data):

    pos_rel = dpg.get_value(dpg_move_pos)
    pos_abs = [90-pos_rel]
    
    xps1.move_abs("XY", pos_abs[0])
    rel_pos_str = f"Position: [{str(pos_rel)}]"
    dpg.set_value(dpg_rel_pos_str, rel_pos_str)
    

def nd_wav_callback(sender, data):

    wav = dpg.get_value(dpg_wav_val)
    nd.wavelength=wav
    nd.set_wavelength(wav)

def on_update():

    global live_pow
    global live_time

    #if(dpg.get_frame_count() % 30 == 0):
    #    [actualwavelength, mean_power, std_power] = nd.read_buffer(663, 10, 1)
    #    live_pow += [float(mean_power)]
    #    live_time += [dpg.get_total_time()]
    #    dpg.set_value(live_series, [live_time, live_pow])
    #    print("Time ", dpg.get_total_time())

def add_exp_to_queue_callback():
    global experiment_queue
    
    #Read information from GUI
    exp_range = dpg.get_value(dpg_range)
    step = dpg.get_value(dpg_step)
    directory= dpg.get_value(dpg_dir)
    file = dpg.get_value(dpg_file)   
    
    # estimate duration
    # estimate takes 145s for 201 steps ->   0.72 seconds/step
    degrees_scanned=abs(exp_range[1]-exp_range[0])
    n_steps=int(degrees_scanned/step)+1
    duration=(n_steps*0.72/60)+0.5 #minutes, time added for initial positioning

    # now create element to add to queue
    message=f'Range: {exp_range[0:2]}, Step: {np.round(step,3)}'
    print(f'ADDED TO QUEUE: {message}')
    
    element=QueueElement(message, lambda: experiment(exp_range, step, directory, file), duration)
    experiment_queue+=[element]
    show_queue_callback()

def add_wait_to_queue_callback():
    global experiment_queue
    
    # Read information from GUI
    wait_time=dpg.get_value(dpg_wait_length)
    
    # now create element to add to queue
    message=(f'Wait: {wait_time} seconds')
    print(f'ADDED TO QUEUE: {message}')
    
    element=QueueElement(message, lambda: wait(wait_time), wait_time/60)
    experiment_queue+=[element]
    show_queue_callback()

def clear_queue_callback():
    global experiment_queue
    experiment_queue=[]
    print('QUEUE CLEARED.')
    show_queue_callback()

def copy_element_callback():
    from copy import deepcopy
    global experiment_queue
    
    # Read information from GUI
    copy_index=dpg.get_value(dpg_copy_row)
    
    if copy_index <=len(experiment_queue):
        # copy element with this index
        element=experiment_queue[copy_index]
        experiment_queue.append(deepcopy(element)) 
        show_queue_callback()
    else:
        print("Index too high.")

def delete_element_callback():
    global experiment_queue
    
    # Read information from GUI
    delete_index=dpg.get_value(dpg_copy_row)
    
    # delete element with this index
    if delete_index <=len(experiment_queue):
        experiment_queue.pop(delete_index)
        show_queue_callback()
    else:
        print("Index too high.")    

def run_queue_callback():
    global experiment_queue
    
    # This a blocking function. Will cause problems if we wish to abort the experiment or continue adding to queue.
    
    if len(experiment_queue)>=1:
        for idx, element in enumerate(experiment_queue):
            print(f"Running: {element.name}.")
            element.execute()
            experiment_queue[idx].duration=0.00
            show_queue_callback()

        print("Queue finished!")
        experiment_queue=[]
        show_queue_callback()
    else:
        MainWindow_width = dpg.get_item_width(dpg_main)
        MainWindow_height = dpg.get_item_height(dpg_main)
    
        with dpg.window(label="Warning", modal=True, show=True, tag="warning", no_title_bar=True, pos=[]) as modalwindow:
            dpg.add_text("The queue is empty. First add some jobs!")
            dpg.add_separator()
            with dpg.group(horizontal=True):
                dpg.add_button(label="OK", width=250, callback=lambda: dpg.delete_item('warning', children_only=False)) 
                
        ModalWindow_width = dpg.get_item_width(modalwindow)
        ModalWindow_height = dpg.get_item_height(modalwindow)
        dpg.set_item_pos(modalwindow, [int((MainWindow_width/2 - ModalWindow_width/2)), int((MainWindow_height/2 - ModalWindow_height/2))])

def show_queue_callback():
    global experiment_queue

    dpg.delete_item('QueueTable', children_only=False)
    duration_sum=0

    with dpg.table(label='Queue',tag='QueueTable',parent=group1):                            # Adds the headers
        dpg.add_table_column(label='Index',width_fixed=True)   
        dpg.add_table_column(label='Name')
        dpg.add_table_column(label='Runtime [min]')
        
        for count, element in enumerate(experiment_queue):                   
            with dpg.table_row():
                dpg.add_text(f'{count}')
                dpg.add_text(f'{element.name}')
                dpg.add_text(f'{np.round(element.duration,3)}')
                duration_sum+=element.duration

        # add final 'summary' row
        with dpg.table_row():
            dpg.add_text('')
            dpg.add_text('Total Time Remaining->')
            dpg.add_text(f'{np.round(duration_sum,3)}')

def experiment(exp_range,step,directory,file):
    # Generate filename
    file=os.path.join(directory,file)

    # check that this filename doesn't already exist, and if it does then append a number to its title
    file_counter=1
    temp_filename=file+'_'+str(file_counter)+'.csv'
    
    print(f"Path {temp_filename} exists: {os.path.exists(temp_filename)}")

    while os.path.exists(temp_filename):
        file_counter+=1
        temp_filename=file+'_'+str(file_counter)+'.csv'

    print(f"{temp_filename} chosen")
    steps = int((exp_range[1] - exp_range[0])/step)+1
    rel_range_arr = np.linspace(exp_range[0], exp_range[1], steps)
    abs_range_arr = np.linspace(90-exp_range[0], 90-exp_range[1], steps)

    dpg.set_axis_limits(pow_x, exp_range[0], exp_range[1])

    df_columns=["Angle","Power","Std_power"]
    out=pd.DataFrame(columns=df_columns)

    #out += nd.read_buffer(
    #    wavelength=663, buff_size=10, interval_ms=1)
    position = []
    #position += [0]
    pow = []
    pos = []

    for i in abs_range_arr:
        
        xps1.move_abs("XY", i)
        [mean_power, std_power] = nd.read_buffer()
        out.loc[len(out)]=[float(90-i), float(mean_power), float(std_power)]
        position += [90-i]
        pow += [float(mean_power)]
        pos += [float(90-i)]
        
        dpg.set_value(pow_series, [pos, pow])
        
    position = np.array(position)
    
    out.to_csv(temp_filename,mode='a',header=True)

    #dpg.set_axis_limits(pow_y, np.min(pow), np.max(pow))
    dpg.set_value(pow_series, [pos, pow])

def wait(wait_time):
    time.sleep(wait_time)

# ===================================

if __name__ == '__main__':

    try:
        nd = Newport_2936(interval_ms=1)
        xps1 = Newport_XPS("XY")
        xps1.xps.set_velocity(stage="XY.X",velo=5,accl=4)
        xps1.xps.set_velocity(stage="XY.Y",velo=5,accl=4)
        
        if nd.status == 'Connected':
            print('Serial number is ' + str(nd.serial_number))
            print('Model name is ' + str(nd.model_number))

            # Print the IDN of the newport detector.
            print('Connected to ' + nd.ask('*IDN?'))
            #print(nd.ask('PM:MAX:Lambda?'))    
        else:
            nd.status != 'Connected'
            print('Cannot connect.')
    except:
        print('Problem with connection. Opening as dummy instead.')

        rel_pos = [np.nan]
        rel_pos_str = f"Position: [{str(rel_pos[0])}]"
    

    live_pow = []
    live_time = []

    experiment_queue=[]

    dpg.create_context()
    dpg.create_viewport()
    dpg.setup_dearpygui()

    with dpg.window(label="Example Window") as dpg_main:

        with dpg.theme() as dpg_plot_theme:
            with dpg.theme_component(dpg.mvLineSeries):
                dpg.add_theme_color(dpg.mvPlotCol_Line, (60, 150, 200), category=dpg.mvThemeCat_Plots)
                dpg.add_theme_style(dpg.mvPlotStyleVar_Marker, dpg.mvPlotMarker_Square, category=dpg.mvThemeCat_Plots)
                dpg.add_theme_style(dpg.mvPlotStyleVar_MarkerSize, 4, category=dpg.mvThemeCat_Plots)

        with dpg.group(horizontal=True) as group1:
            with dpg.group(label="Controls"):
                dpg.add_text("Motion Controls")
                dpg_rel_pos_str = dpg.add_text(rel_pos_str)
                with dpg.group(horizontal=True):
                    dpg_home = dpg.add_button(label="Home Laser & Detector", callback=xps_home_callback, width=200)
                
                dpg_move_pos= dpg.add_input_int(default_value=80,max_value=90, min_value=30, max_clamped=True, min_clamped=True, width=200)

                with dpg.group(horizontal=True):
                    dpg_move = dpg.add_button(label="Move Laser & Detector", callback=xps_move_callback, enabled=False, width=200)

                dpg.add_text("Power Meter Controls")
                dpg_wav_val = dpg.add_input_int(default_value=633, width=200)
                dpg_wav = dpg.add_button(label="Set Wavelength", callback=nd_wav_callback, width=200)
                
                dpg.add_text("Experiment Controls")
                dpg_range = dpg.add_input_intx(label="Range", size=2, default_value=[30, 60], max_value=90, min_value=30, max_clamped=True, min_clamped=True, width=200)
                dpg_step = dpg.add_input_float(label="Precision", width=200, default_value=0.10)
                
                dpg.add_text("Save Location:")

                dpg_dir = dpg.add_input_text(label="Save directory", default_value=os.path.abspath(os.curdir), width=200)
                dpg_file = dpg.add_input_text(label="File Name", default_value="experiment", width=200)

                dpg.add_text("")
                dpg.add_button(label="Queue Experimental Run",callback=add_exp_to_queue_callback,width=200)
                
                dpg.add_text("")
                dpg.add_text("Wait timer (seconds):")
                dpg_wait_length = dpg.add_input_int(default_value=300, width=200)
                dpg.add_button(label="Queue Wait",callback=add_wait_to_queue_callback,width=200)
                
                dpg.add_text("")
                dpg.add_text("Queue manipulation:")
                dpg_copy_row = dpg.add_input_int(label="Index",default_value=0, width=200)
                with dpg.group(horizontal=True):
                    dpg.add_button(label="Copy Row",callback=copy_element_callback,width=100)
                    dpg.add_button(label="Delete Row",callback=delete_element_callback,width=100)
            
                dpg.add_button(label="Clear Queue", callback=clear_queue_callback, width=200)
                dpg.add_text("")
                
                dpg_run=dpg.add_button(label="Run Queue", callback=run_queue_callback,enabled=False, width=200)
                
            with dpg.plot(label="Experiment", height=700, width=700) as dpg_plot:
                pow_x = dpg.add_plot_axis(dpg.mvXAxis, label="Angle")
                pow_y = dpg.add_plot_axis(dpg.mvYAxis, label="Power [W]")
                pow_series = dpg.add_line_series([], [], parent=pow_y)
                dpg.set_axis_limits(pow_x, 30, 60)
                dpg.bind_item_theme(dpg_plot, dpg_plot_theme)
                
            with dpg.table(label='Queue',tag='QueueTable'):                            # Adds the headers
                dpg.add_table_column(label='Index',width_fixed=True)   
                dpg.add_table_column(label='Name')
                dpg.add_table_column(label='Runtime')
                
                # Add a nonsense first row
                
                with dpg.table_row():
                    dpg.add_text('-')
                    dpg.add_text('-')
                    dpg.add_text('-')

                    # for count, element in enumerate(experiment_queue):                   
                    #     with dpg.table_row():
                    #         dpg.add_text(f'{count}')
                    #         dpg.add_text(f'{element.name}')
                    #         dpg.add_text('None')

            #with dpg.plot(label="Live Power", height=800, width=800) as dpg_live:
            #    live_x = dpg.add_plot_axis(dpg.mvXAxis, label="Time")
            #    live_y = dpg.add_plot_axis(dpg.mvYAxis, label="Power")
            #    #dpg.set_axis_limits(pow_x, 30, 60)
            #    dpg.bind_item_theme(dpg_plot, dpg_plot_theme)

    dpg.show_viewport()
    dpg.maximize_viewport()
    dpg.set_primary_window(dpg_main, True)

    while dpg.is_dearpygui_running():
        dpg.render_dearpygui_frame()
        #on_update()

    dpg.destroy_context()


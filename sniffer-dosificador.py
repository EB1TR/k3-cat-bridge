
import serial
import time

fa_data = ""
if_mode = 0
fa_time = 0
fa_last_send = 0
fa_send_interval = 0.1
fa_expire = 0.25

OmniRig = False


def send_serial():
    global fa_data
    global fa_time
    global fa_last_send
    serb.write(fa_data.encode('utf-8'))
    fa_last_send = time.time()
    print("RS232 | %s | %s | %s" % (msg, int(fa_last_send), fa_data))


if OmniRig:
    import win32com.client
    rig = win32com.client.Dispatch("OmniRig.OmniRigX")
    try:
        serb = serial.Serial(port='COM2')
        serb.baudrate = 9600
        serb.bytesize = serial.EIGHTBITS
        serb.parity = serial.PARITY_NONE
        serb.stopbits = serial.STOPBITS_TWO
        serb.setRTS = False
        serb.setDTR = False
        serb.rts = False
        serb.dtr = False
        serb.rtscts = False
        serb.dsrdtr = False
    except Exception as e:
        print(e)
        exit(0)
    while True:
        msg = "Stored DATA "
        # Determine if the program need send or send and poll data
        time.sleep(0.05)
        fa_send = fa_last_send + fa_send_interval <= time.time()
        fa_expired = fa_time + fa_expire <= time.time()
        if fa_expired:
            msg = "OmniRig DATA"
            data_a = "0000%s" % rig.Rig1.Freq
            fa_data = "FA%s;" % data_a[-11:]
            fa_time = time.time()
        if fa_send:
            send_serial()
else:
    try:
        sera = serial.Serial(port='COM4')
        sera.baudrate = 38400
        sera.bytesize = serial.EIGHTBITS
        sera.parity = serial.PARITY_NONE
        sera.stopbits = serial.STOPBITS_TWO
        sera.setRTS = False
        sera.setDTR = False
        sera.rts = False
        sera.dtr = False
        sera.rtscts = False
        sera.dsrdtr = False
        serb = serial.Serial(port='COM2')
        serb.baudrate = 9600
    except Exception as e:
        print(e)
        exit(0)

    while True:
        try:
            # Determine if the program need send or send and poll data
            time.sleep(0.05)
            fa_send = fa_last_send + fa_send_interval <= time.time()
            fa_expired = fa_time + fa_expire <= time.time()

            # Sniffer of the PC<->Rig
            if sera.inWaiting() > 0:
                data_a = sera.read_until(b';')
                data_a = data_a.decode('utf-8')
                if data_a[:2] == "FA":
                    fa_data = "FA%s;" % data_a[2:13]
                    fa_time = time.time()
                elif data_a[:2] == "IF":
                    fa_data = "FA%s;" % data_a[2:13]
                    fa_time = time.time()
                    if_mode = data_a[29:30]

            # Send stored data or poll, if is necessary
            if fa_send:
                # Send stored data
                if not fa_expired:
                    msg = "Stored DATA "
                    send_serial()
                # Poll, store and send data
                else:
                    msg = "Polled DATA "
                    sera.write(b'FA;')
                    data_a = sera.read_until(b';')
                    data_a = data_a.decode('utf-8')
                    if data_a[:2] == "FA":
                        fa_data = "FA%s;" % data_a[2:13]
                        fa_time = time.time()
                        send_serial()

        except Exception as e:
            print(e)

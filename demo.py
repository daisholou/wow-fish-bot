
import time
import sounddevice as sd
import pyaudio
import audioop
import math
from collections import deque
import wave

def listen():
    print('Well, now we are listening for loud sounds...')
    CHUNK = 1024  # CHUNKS of bytes to read each time from mic
    FORMAT = pyaudio.paInt16
    CHANNELS = 2
    RATE = 18000
    THRESHOLD = 1000  # The threshold intensity that defines silence
    # and noise signal (an int. lower than THRESHOLD is silence).
    SILENCE_LIMIT = 1  # Silence limit in seconds. The max ammount of seconds where
    # only silence is recorded. When this time passes the
    # recording finishes and the file is delivered.
    # Open stream
    p = pyaudio.PyAudio()

    stream = p.open(format=FORMAT,
                    channels=CHANNELS,
                    rate=RATE,
                    input=True,
                    input_device_index=2,
                    frames_per_buffer=CHUNK)

    # stream = p.open(format=FORMAT,
    #                 channels=CHANNELS,
    #                 rate=RATE,
    #                 input=True,
    #                 frames_per_buffer=CHUNK)
    cur_data = ''  # current chunk  of audio data
    rel = RATE / CHUNK
    print(rel)
    slid_win = deque(maxlen=SILENCE_LIMIT * int(rel))

    success = False
    listening_start_time = time.time()
    record_buf = []
    while True:
        cur_data = stream.read(CHUNK)
        # print(cur_data)
        record_buf.append(cur_data)
        slid_win.append(math.sqrt(abs(audioop.avg(cur_data, 4))))
        print(time.time() - listening_start_time)
        print(sum([x > THRESHOLD for x in slid_win]))
        # if sum([x > THRESHOLD for x in slid_win]) > 0:
        #     print('I heart something!')
        #     success = True
        #     break
        if time.time() - listening_start_time > 20:
            print('I don\'t hear anything already 20 seconds!')
            break


    # print "* Done recording: " + str(time.time() - start)

    stream.close()
    p.terminate()
    return success


if __name__ == "__main__":
     listen()

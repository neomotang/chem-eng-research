"""
< -~\|/~-  Welcome to Amethystos Viscosity App v2.0  -~\|/~- >

By: Neo Motang <neomotang@gmail.com>, based on software by:
      Herman Franken <hhfranken@sun.ac.za>
      Mark Harfouche < mark.harfouche@gmail.com>

It uses Siggen with the PicoScope 2000A
It was tested with the PS2206B USB2.0 version
The SigGen is connected to Channel A
See http://www.picotech.com/ for PS2000A models
See github.com/colinoflynn/pico-python for driver specific files

"global ps" might give a warning; program works anyway :P
			
"""

from __future__ import division
from __future__ import print_function
from __future__ import unicode_literals
import time
#from picobase import ps2000a #For other PS models, change "ps2000a" to required model
import ps2000a
import numpy as np

# Function for writing to a text file:
def txt_write(FILENAME, DATA, MODE):
    f = open(FILENAME, MODE)
    f.write(DATA + "\n")
    f.close()
    return()
    
# Function for setting Frequency parameters for the PicoScope	
def freqset ():
	
	m = -1	# Initialize m counter - menu counter

	#defaults
	start_freq = 39000
	end_freq = 39200
	increment = 10
	timeleng = 1
	
	while m != 0:
	
		# Menu Text
		print("\n\n")
		print("Frequency selection menu: \n")
		print("Select the frequency scan mode:")
		print("1: Define Frequency Start, End and Increment")
		print("2: Use defaults: 39000 - 39200 Hz, increment of 10 Hz")
		print("0: Exit")
		
		# Menu Input
		try:
			m = int(input("option: "))
		except:
			print("\n Invalid input - try again")	
	
	# m validity check
		if m > 3 or m < -1:
			print("\n Invalid input - try again")
			
		# Menu opt 1
		if m == 1:
		
			start_freq = float(input("Start Frequency (Hz): "))
			end_freq = float(input("End Frequency (Hz): "))
			increment = float(input("Increment (Hz): "))
			
			timeleng = ((end_freq - start_freq) / increment) * 3.175  # constant to reflect true processing time. 
			
			print("\n\n")
			print("Frequencies selected:")
			print("Start:",start_freq)
			print("End:",end_freq)
			print("Increment:",increment)
			print("Estimated runtime of scan (s):",timeleng)
			
			break

		# Menu opt 2
		if m == 2:
		
			timeleng = ((end_freq - start_freq) / increment) * 3.175  # constant to reflect true processing time. 

			print("\n\n")
			print("Defaults selected:")
			print("Start:",start_freq)
			print("End:",end_freq)
			print("Increment:",increment)
			print("Estimated runtime of scan (s):",timeleng)
			
			break
		
		# Hidden Menu opt 123 - skip
		if m == 123:
			break
		
		# Menu opt 0
		if m == 0:
			return()
			
		m = -1	# Reset m
	return(start_freq,end_freq,increment,timeleng)

# Main Function
def main():
		
	print(__doc__)

	m = -1	# Initialize m counter - menu counter
	 
	### MENU 1 - Confirm PicoScope connected and ready

	while m != 0:
		
		# Menu text
		print("\n")
		print("Is the PicoScope connected: - Connect to USB FIRST \n")
		print("1: Yes")
		print("2: No")
		print("3: Help")
		print("0: Exit")
		
		# Menu Input
		try:
			m = int(input("option: "))
		except:
			print("\n Invalid input - try again")

		# m validity check
		if m > 4 or m < -1:
			print("\n Invalid input - try again")
			
		# Menu opt 1
		if m == 1:
			#ps = ps2000a.PS2000a(connect=False) #closes any existing connection to scope
			print("\n")
#			print("Attempting to connect to PicoScope 2000A...")	
			global ps
			ps = ps2000a.PS2000a()
			print("Successfully connected to PicoScope model " + ps.getUnitInfo('VariantInfo'))  
			#print(ps.getAllUnitInfo())
			
			break
		
		# Menu opt 2
		if m == 2:
			print("\n")
			print("Please connect PicoScope and try again")	
			m = -1
			
		# Menu opt 3
		if m == 3:
			print("\n\n")
			print("HELP:")
			print("\n")
			print("For PicoScope help, go to http://www.picotech.com/")
			print("For python files help, go to github.com/colinoflynn/pico-python")
			print("NB - make sure PicoScope SDK is installed and that the scope is connected FIRST")
			m = -1
		
		# Hidden Menu opt 123 - skip
		if m == 123:
			break
		
		# Menu opt 0
		if m == 0:
			return()
			
		m = -1	# Reset m	

	### MENU 2 - Set frequency scan values

	start_freq,end_freq,increment,timeleng =	freqset()
		
	filename = str("test.csv")	#Default filename
	
	### MENU 3 - Run Scan
		
	while m != 0:	
			
		# Menu Text
		print("\n\n")
		print("Start testing:\n")
		print("1: Start Test")
		print("2: Change frequency parameters")
		print("3: Input Filename (default is 'test')")
		print("0: Exit")	
			
		# Menu Input
		try:
			m = int(input("option: "))
		except:
			print("\n Invalid input - try again")	
	
		# m validity check
		if m > 4 or m < -1:
			print("\n Invalid input - try again")
	
		# Menu opt 1
		if m == 1:

			#Write time measurement parameters to file
			localtime = time.asctime( time.localtime(time.time()) )
			futuretime = time.asctime( time.localtime(time.time()+ timeleng) )
			print ("Run starting at: ", localtime)
			print("\n")
			print ("Approximate end: ", futuretime)
			print("\n")
			txt_write(filename, "Data collection started at: " + localtime, 'a')
						
			# Set oscilloscope channel and signal generator trigger
			ps.setChannel('A', 'DC', 1.0, 0.0, True, False, 10.0)
			ps.setSimpleTrigger('A', 0.0, 'Rising', delay=0, timeout_ms=100, enabled=True)

			obs_duration = 1  #frequency hold/observation time in s
			sampling_interval = obs_duration / 500000 #based on 500000 samples

			(actualSamplingInterval, nSamples, maxSamples) = ps.setSamplingInterval(
			sampling_interval, obs_duration)

			# Set and sweep frequency
			freq = start_freq
			while freq <= end_freq:
				ps.setSigGenBuiltInSimple(offsetVoltage=0, pkToPk=2, waveType="Sine", frequency=freq,shots=1, triggerType="Rising",triggerSource="None")
				
				ps.runBlock()
				ps.waitReady()
				time.sleep(1)
				ps.runBlock()
				ps.waitReady()
				print("     %f Hz" % (freq))
				dataA = ps.getDataV('A', nSamples, returnOverflow=False) #Records timebase measurements into an array
				txt_write(filename, str(freq) + " " + str(1000*np.sqrt(np.mean(dataA**2))), 'a')	#Write measurements to file	
				freq = freq + increment		#Freqency Step
   			
		# Menu opt 2
		if m == 2:
		
			start_freq,end_freq,increment,timeleng =	freqset()
			
		# Menu opt 3
		if m == 3:
		
			filename = str (input("Input Filename (start & end with '): "))
			filename = filename + ".csv"
			
			print("\n\n")
			print("Filename selected:",filename)
		
		# Hidden Menu opt 123 - skip
		if m == 123:
			break
		
		# Menu opt 0
		if m == 0:
			return()
			
		m = -1	# Reset m

# Start of code:
if __name__ == "__main__":
	# Call main function:
	main()
	# Exit confirmation:
	print ("\nTest done")
	ps.close()
	#pass

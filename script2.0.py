

## STEP 1: Import Liraries needed to run the code
import win32com.client as com
import os

### STEP 2: Connect to COM server and open a new Vissim Window
## Connecting the COM Server => Open a new Vissim Window:
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")


### STEP 3: Get address of the directory where Vissim file is.
### This is not a necessary step. It is used to let the script know where to find the file you want to run
### The address to folder from where this script is running is stored in Path_of_COM_Basic_Commands_network
Path_of_COM_Basic_Commands_network = os.getcwd()

## You can print out and check if the command os.getwd(), which means - operating system.get working directory gives you the correct address
print (os.getcwd())

## STEP 3: Add name of the Vissim file to the address we fetched in step 3. Store complete address in Filename
Filename = os.path.join(Path_of_COM_Basic_Commands_network, 'Module 6 031521.inpx')

## STEP 4: Load Vissim inp file on Vissim instance using function Vissim.LoadNet(address to file)
Vissim.LoadNet(Filename)

## STEP 5: Set the values of parameters CC0, CC1,CC2 that need to be varied in each simulation run 
# CC0 values to be varied
ccZero_list = [5,6,7,8,9]
# CC1 values to be varied
ccOne_list = [1,2,3,20,30]
# CC2 values to be varied
ccTwo_list = [5,6,7,8,9]
# CC3 values to be varied
ccThree_list = [5,6,7,8,9]

## List of CCs being varied
cc_list = [ccZero_list, ccOne_list, ccTwo_list, ccThree_list]

## STEP 6: For loops to run simulation for each of the values in the CC0, CC1, CC2, and CC3
## Following code will run one simulation for each of the CC0, CC1, CC2, and CC3 in the list, with other parameters being constant
for j in range(len(cc_list)):
	for k in range(len(cc_list[j])):

	    ## STEP 7: Set Simulation Run Attributes
	    ## Set Simulation time. Vissim.Simulation.SetAttValue can e used to access and set several Vissim Simulation Run attriutes
	    simtime = 10

	    ## Set simulation time
	    Vissim.Simulation.SetAttValue('SimPeriod', simtime)

	    ## Set simulation to use max simulation speed
	    Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)

	    ## Set random seed for the simulation run
	    #Vissim.Simulation.SetAttValue('RandSeed', Random_Seed[r])        

	    #Get Simulation Resolution attriute from simulation model set in Vissim interfacce
	    simRes = Vissim.Simulation.AttValue('SimRes')

	    #initiate variale i that represents simulation step 
	    i = 0

		#create a while loop to run simulation that loops over variable i
	    while (i<=((simtime-1)*simRes)):

	    	#At start of simulation when i is 0, the values of the weidemann parameters will be changed. 
	    	#This can be changed to some other smulation time, 
	    	#if the parameter value changes are to be made in the middle of the simulation

	    	if (i==0):

	    		## STEP 8: j is 0 implying CC0 values are changed. 
	    		## Change the values for the CC1, CC2, and CC3 as required. 
	    		## These will be constant and CC0 Values will be changing as per the for loop
		        if (j==0):
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc0', cc_list[j][k])
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc1Distr', 1)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc2', 2)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc3', 4)
		            print(j)
		            print(k)
		            print(cc_list[j][k])

		        ## STEP 9: j is 1 implying CC1 values are changed. 
	    		## Change the values for the CC0, CC2, and CC3 as required. 
	    		## These will be constant and CC1 Values will be changing as per the for loop 
		        elif (j==1):
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc1Distr', cc_list[j][k])
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc0', 1)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc2', 2)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc3', 4)
		            print(j)
		            print(k)
		            print(cc_list[j][k])

		        ## STEP 10: j is 2 implying CC2 values are changed. 
	    		## Change the values for the CC0, CC1, and CC3 as required. 
	    		## These will be constant and CC2 Values will be changing as per the for loop
		        elif (j==2):
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc2', cc_list[j][k])
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc0', 1)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc1Distr', 2)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc3', 4)
		            print(j)
		            print(k)
		            print(cc_list[j][k])


		        ## STEP 11: The last possible value of j is 3 implying CC3 values are changed. 
	    		## Change the values for the CC0, CC1, and CC2 as required. 
	    		## These will be constant and CC3 Values will be changing as per the for loop
		        else:
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc2', cc_list[j][k])
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc0', 1)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc1Distr', 2)
		            Vissim.Net.DrivingBehaviors.ItemByKey(104).SetAttValue('W99cc3', 4)
		            print(j)
		            print(k)
		            print(cc_list[j][k])


    		## STEP 10: Run Simulation Step and Increment Value of i
	    	Vissim.Simulation.RunSingleStep()
	    	i=i+1

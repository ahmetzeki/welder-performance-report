### welder-performance-report

Problem: High ratio of rejected joints in construction of Gas processing plant project
 
#  1- Creation of ndt-result database to calculate welder performance: 
   I obtain data about the joints the welder worked on from files called 'New Joints' and 'Repair Joints.' Next, I retrieve data from the 'Ndt-Status' file to determine the status of each joint (whether it was accepted or not). Additionally, I collect information on the type of defects made by the welder during the welding process. Then, from the 'Based-Tp' file, I extract joint characteristics, including type, thickness, inch, size, and welding method. As a result, I have all the parameters of the joints where the welder worked and the type of defects they made during the welding process.

#  2- Welder performance generator: 
   Once I have all the data ready, I calculate the percentage of accepted joints for each welder based on the joint's parameters. This information helps us make decisions about which type of joint each welder should work on, leading to increased effectiveness. By analyzing the acceptance rates of different joint types for each welder, we can identify their strengths and assign them to tasks that match their skills, ultimately optimizing the welding process and improving overall efficiency.

# Libraries

   Pandas,
   Numpy,
   Openpyxl

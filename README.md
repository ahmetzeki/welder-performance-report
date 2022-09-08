### welder-performance-report

Problem: High ratio of rejected joints in construction of Gas processing plant project
 
#  1- Creation of ndt-result database to calculate welder performance: 
   I get data about which joints the welder worked on from files ('New Joints' and 'Repair Joints'), next I get data from Ndt-Status file to see the status of the joint(if accepted or not), also I get type of defect data to see what kind of defect the welder made during the process of welding. Then, from Based-Tp file, I pull characteristics of joints such as type, thicness, inch, size and welding method. As a result i have all parameters of joint where the welder worked on and what kind of defects he made

#  2- Welder performance generator: 
   Once I have all data ready, I calculate the persentage of accepted joints for each welder according to joint's parameters. This info help us make decision which type of joint sould a welder work on and we increase effectivity this way.

# Libraries

   Pandas,
   Numpy,
   Openpyxl

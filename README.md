# Excel-Cell-Analysis
A script with the direct goal of automating sequential operations to lines on an excel spreadsheet using the OpenPyxl library in python. Note: even tough the script can be used for different purposes, it was mainly designed for a specific task in my workload, altough many parts of it can be reutilized flawlessly with previous analysis of the type of iteration needed  



# Function and Class detailing:
 ## Participante is an arbitrary class used to store two variables, **the time spend in the meeting** (tempo) **and the name of the participant** (nome)
  
 ## _calculata_tempo(string)_ initially parses the string, transforming it into an array of its substrings, then proceeds to iterate over the string to calculate the time spend in seconds(it does so using simple logic, derivated from the document's format)
  
 ## **since the program relies in a specific format**, _reformatting is necessary for other types of documents / use-cases._
 ## formata_tempo(tempo) simply takes the input time in seconds (tempo) and puts it in a "xx h xx min xx sec" format to be inserted into a cell
 
 ## into the main loop, since the file was successfully loaded (the "file", "pasta" and "nome_pasta"variables are reserved for that), the main loop will iterate trough the lines in the spreadsheet, starting in the position x and ending when it encounters a empty cell, it will execute the functions listed above and if the cell format is correct in all of them, itwill print the total time in a new sheet called "Tempo", with the respective names of the participants and their total time spent
  

if dataFull.split("\n")[3].strip() == "Results" and dataFull.split("\n")[6].strip() == "=======================================================================":
                                        print("Type 6 Exists for " + entry)
                                        dataArray = dataFull.split("=======================================================================")
                                        data = []
                                        nextGood = False
                                        for dataChunk in dataArray:
                                            if nextGood == True:
                                                nextGood = False
                                                dataSplits = dataChunk.split("\n")
                                                #print(dataSplits[0])
                                                #print(dataSplits[1])
                                                #print(dataSplits[2])
                                                for dataLine in dataSplits:
                                                    try:
                                                        x = int(dataLine[0:4:].strip())
                                                        #print("-" + str(x) + "-")
                                                        data.append(dataLine)
                                                    except:
                                                        boolean = False
                                            if dataChunk.strip().startswith("Name") or dataChunk.strip()[0:4] == "Name":
                                                nextGood = True
                                                #print("Aaaaaaaaaaaaaaaaaaa")
                                            
                                        #print(str(data))
                                        for personString in data:
                                            newPerson = Person()
                                            names = personString[4:30:].strip().split(" ")
                                            #print(personString)
                                            #print(names)
                                            newPerson.firstName = names[0]
                                            try:
                                                newPerson.lastName = names[1]
                                            except:
                                                boolean = False
                                            newPerson.school = personString[33:56:].strip()
                                            newPerson.time = personString[56:63:].strip()
                                            
                                            gradeRude = personString[30:33].strip()
                                            if gradeRude != "":
                                                newPerson.grade = grades[gradeRude]

                                            #print(newPerson.firstName + " " + newPerson.lastName + " (grade " + str(newPerson.grade) + "): " + newPerson.school + ", " + newPerson.time)
                                            currentRace.people.append(newPerson)
                                            if newPerson.firstName == "Sebastian" and newPerson.lastName == "Polge":
                                                print("added " + newPerson.firstName + " " + newPerson.lastName)
    
                                        raceInfo = dataFull.split("\n")
                                        currentRace.name = raceInfo[0].strip().split(" - ")[0]
                                        currentRace.courseName = raceInfo[2].strip()
                                        raceDateString = raceInfo[0].strip().split(" - ")[1].strip().split(" to ")[0].strip().split("/")
                                        currentRace.date = datetime.date(int(raceDateString[2]), int(raceDateString[0]), int(raceDateString[1]))
                                        courseData[currentRace.name] = currentRace
                                        #print(currentRace.name + ": hosted at " + currentRace.courseName + " on " + str(currentRace.date))

                                    else:
                                        
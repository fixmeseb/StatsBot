newRacer = Racer()
        if name in people.keys():
            newRacer = people[name]
            racesParticipated = ""
            
            racesParticipated = racesParticipated + " and " + courseName
            #print(courseName + "\t\tFOUND " + name + " in " + racesParticipated + "!")
        try:
            if name == "Sebastian Polge":
                print(name + ": " + person.time)
            newRacer.times[courseName] = person.time
        except:
            newRacer.times[courseName] = "NULL"
        newRacer.firstName = person.firstName
        newRacer.lastName = person.lastName
        #if newRacer.school != "" and newRacer.school != person.school:
            #print(courseName + "\t" + newRacer.firstName + " " + newRacer.lastName + " ran for " + newRacer.school + " AND " + person.school)
        newRacer.school = person.school
        people[name] = newRacer
        if name == "Sebastian Polge":
            print(name + ": " + str(newRacer.times.values()))




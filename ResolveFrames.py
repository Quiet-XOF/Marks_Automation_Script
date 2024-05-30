import csv
import re
import sys


def resolveFiles(files):
    try:
        if len(files) < 3:
            print("Not enough files given.")
            sys.exit()
        file = open(files[1])
        order = file.readlines()
        file.close()
        if order[0].startswith("/"):
            print("Please give workorder file first.")
            sys.exit()
        work = []
        for i in range(len(files)-2):
            file = open(files[i+2])
            work += file.readlines()
            file.close()
        return order, work
    except:
        print("How to use: Project1.py workorder.txt work1.txt work2.txt ...")
        sys.exit()


def getData(order, work, dictionary):
    try:
        dictionary['workorder'] = next(line.strip() for line in order if 'workorder' in line.lower())
        dictionary['producer'] = next(line.split(":")[1].strip() for line in order if 'producer' in line.lower())
        dictionary['operator'] = next(line.split(":")[1].strip() for line in order if 'operator' in line.lower())
        dictionary['job'] = next(line.split(":")[1].strip() for line in order if 'job' in line.lower())
        dictionary['notes'] = order[next(index + 1 for index, line in enumerate(order) if 'notes' in line.lower())]
        dictionary['location'] = [line.strip() for line in order if line.startswith('/')]
        dictionary['work'] = [line.strip() for line in work if line.strip()]
    except:
        print("There was an error in the work order file.")
        sys.exit()
    return dictionary


def cleanUp(dictionary):
    try:
        for i, line in enumerate(dictionary['work']):
            line = re.sub(r'<[^>]+>', '', line).strip() # remove errors
            line = '/'.join(line.split('/')[2:]) # remove old location
            local, frames = line.split(' ', 1)
            line = next(base for base in dictionary['location'] if local in base)
            frames = list(map(int, frames.split())) # make frames ints
            start = current = frames[0]
            for frame in frames[1:]:
                if frame == current + 1: current = frame
                else:
                    if start == current: dictionary['matched'].append(f"{line} {start}")
                    else: dictionary['matched'].append(f"{line} {start}-{current}")
                    start = current = frame
            if start == current: dictionary['matched'].append(f"{line} {start}")
            else: dictionary['matched'].append(f"{line} {start}-{current}")
    except:
        print("There was an error with frame cleanup.")
        sys.exit()
    return dictionary


def writeCSV(dictionary):
    try:
        with open(f"{dictionary['workorder']}.csv", 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)   
            # Line 1: Producer / Operator / Job / Notes
            writer.writerow(["Producer", "Operator", "Job", "Notes"])
            writer.writerow([dictionary['producer'], dictionary['operator'], dictionary['job'], dictionary['notes']])
            # Line 4: Show Location / Frames to Fix
            writer.writerow('') # skip line
            writer.writerow(["Show Location", "Frames to Fix"])
            for line in dictionary['matched']:
                local, frame = line.split(' ', 1)
                writer.writerow([local, f"{frame}"])
            csvfile.close()
    except:
        print("There was an error writing to the csv.")
        sys.exit()


def main():
    dictionary = {'workorder': None, 'producer': None, 'operator': None, 'job': None, 
    'notes': None, 'location': None, 'work': None, 'matched': []}
    order, work = resolveFiles(sys.argv)
    print("Files accepted.")
    dictionary = getData(order, work, dictionary)
    print("Cleaning...")
    cleanUp(dictionary)
    print("Writing to file...")
    writeCSV(dictionary)
    print("Finished! Program ending...")


if __name__ == "__main__": main()

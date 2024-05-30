import argparse
import csv
import datetime
import ffmpeg
from frameioclient import FrameioClient
from io import BytesIO
import os
import pandas
import pymongo
import shlex
import subprocess

from ResolveFrames import getData, cleanUp


def parse_args():
    parser = argparse.ArgumentParser(description="Everything hurts and I'm dying.")
    # Upload information
    parser.add_argument("--baselight", help="Send frames", action="store")
    parser.add_argument("--xytech", help="Send location", action="store")
    parser.add_argument("--process", help="Send video file", action="store")
    parser.add_argument("--output", help="Print to .XLS", action="store")
    return parser.parse_args()


def readFile(file):
    with open(file, "r") as readfile:
        lines = [line.rstrip() for line in readfile.readlines() if line.strip()]
    return lines
    

def createTimecode(frames, framerate):
    seconds = frames/framerate
    hh = int(seconds // 3600)
    mm = int((seconds % 3600) // 60)
    ss = int(seconds % 60)
    ff = int((frames % framerate))
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(hh, mm, ss, ff)


def main():
    args = parse_args()
    client = pymongo.MongoClient(["mongodb://localhost:27017/"])
    db = client["TheCrucible"]
    baselight_col = db["baselight"] # Send folder/frames
    xytech_col = db["xytech"] # Send workorder/location
    process_col = db["processed"] # Organize location, frames and timecodes
    query = {} # For stacked calls
    dictionary = {'workorder': None, 'producer': None, 'operator': None, 'job': None, 
    'notes': None, 'location': None, 'work': None, 'matched': []}

    try: # Check server connection
        client.server_info()
    except pymongo.errors.ConnectionFailure:
        print(" /\\_/\\\n(='w'=)\num i cant find da server")
        sys.exit()

    if args.baselight: 
        lines = readFile(args.baselight)
        for line in lines:
            if baselight_col.find_one({"content": line.strip()}) is None:
                baselight_col.insert_one({"content": line.strip()})

    if args.xytech: 
        lines = readFile(args.xytech)
        for line in lines:
            if xytech_col.find_one({"content": line.strip()}) is None:
                xytech_col.insert_one({"content": line.strip()})

    if args.process:
        framerate = 60
        probe = ffmpeg.probe(args.process)
        duration = float(probe["streams"][0]["duration"])
        last_frame = int(duration * framerate)
        try: # Check if data exists in collection
            baselight = [line["content"] for line in baselight_col.find({}, {"_id": 0})]
            xytech = [line["content"] for line in xytech_col.find({}, {"_id": 0})]
        except:
            print("Please send the correct documents to server.")
            sys.exit()
        # clean up the location and frames
        dictionary = getData(xytech, baselight, dictionary)
        cleanUp(dictionary)
        # convert frames to timecode
        processed = []
        for line in dictionary["matched"]:
            location, frame = line.split(" ", 1)
            if "-" in frame: 
                first, last = frame.split("-", 1)
                if int(first) > last_frame or int(last) > last_frame: break
                timecode_start = createTimecode(int(first), framerate)
                timecode_last = createTimecode(int(last), framerate)
                processed.append(f"{location} {frame} {timecode_start}-{timecode_last}")       
        #for line in processed: print(line)
        for line in processed:
            if process_col.find_one({"content": line.strip()}) is None:
                process_col.insert_one({"content": line.strip()})

    if args.output:
        framerate = 60
        stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        try:
            processed = [line["content"] for line in process_col.find({}, {"_id": 0})]
        except:
            print("Please send the correct documents to server.")
            sys.exit()
        #for line in processed: print(line)
        writer = pandas.ExcelWriter("Work.xlsx", engine="xlsxwriter")
        dictionary = {
            "Location": [],
            "Frame": [],
            "Timecodes": [],
            "Targets": [],
            "Thumbnails": []
        }
        for line in processed:
            local, frame, timecode = line.split(" ", 2)
            dictionary["Location"].append(local)
            dictionary["Frame"].append(frame)
            dictionary["Timecodes"].append(timecode)
            first, last = frame.split("-", 1)
            target = (int(first) + int(last)) / 2
            #target = createTimecode(int(target), framerate)
            target = int(target)
            dictionary["Targets"].append(target)
        df = pandas.DataFrame({
            "Show Location": dictionary["Location"],
            "Frames to Fix": dictionary["Frame"],
            "Timecodes": dictionary["Timecodes"],
            "Thumbnails": None
        })
        df.to_excel(writer, index=False)
        workbook = writer.book 
        worksheet = writer.sheets["Sheet1"]
        for i, line in enumerate(dictionary["Targets"]):
            filename = f"thumbnail_{i+1}.png"
            if not os.path.exists(filename):
                command = f"ffmpeg -loglevel 0 -i {args.output} -ss {line/framerate} -vframes 1 -f image2pipe -"
                process = subprocess.Popen(shlex.split(command), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                thumbnail, _ = process.communicate()
                with open(filename, "wb") as file: 
                    file.write(thumbnail)
        for i in range(len(processed)):
            filename = f"thumbnail_{i+1}.png"
            worksheet.insert_image(f"D{i+2}", filename, {"x_scale": 0.07, "y_scale": 0.07})
        #for i in range(len(processed)):
        #    filename = f"thumbnail_{i}.png"
        #    os.remove(filename)
        worksheet.autofit()
        writer._save()
        
        client = FrameioClient("<INSERT TOKEN>")
        project_id = "insert project id"
        folder_id = "insert folder id"
        for i, line in enumerate(dictionary["Frame"]): 
            start, end = line.split("-")
            start = int(start)
            end = int(end)
            start = start / framerate
            end = end / framerate
            end = end - start + (1/framerate)
            #print(start, end)
            command = f"ffmpeg -loglevel 0 -i {args.output} -ss {start} -t {end} clip_{i+1}.mp4"
            process = subprocess.Popen(shlex.split(command), stdout=subprocess.PIPE)
            output, _ = process.communicate()
            process.wait()
            #with open("temp.mp4", "wb") as file:
            #    file.write(output)
            asset = client.assets.upload(folder_id, f"clip_{i+1}.mp4")


if __name__ == "__main__": main()

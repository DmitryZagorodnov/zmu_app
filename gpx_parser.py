import gpxpy
from enums import Places, Animals


class GpxParser:

    def __init__(self):
        self.ans = [animal.value for animal in Animals]
        self.dict_ans = {animal.value: [0, 0, 0] for animal in Animals}
        self.pcs = [place.value for place in Places]
        self.pcs_ans = {place.value: [0, 0, 0] for place in Places}
        self.offset = {
            Places.LES.value: 0,
            Places.POLE.value: 1,
            Places.BOLOTO.value: 2
        }
        self.lats = []
        self.longs = []
        self.track = []
        self.center = []
        self.waypoints = []

    def parse_track(self, tracksfile):
        with open(tracksfile, 'r') as gpx_file:
            gpx = gpxpy.parse(gpx_file)
            self.center = [
                (gpx.bounds.max_latitude + gpx.bounds.min_latitude) / 2,
                (gpx.bounds.max_longitude + gpx.bounds.min_longitude) / 2
            ]
            if not gpx.tracks:
                raise ValueError("File doesn't contains any tracks")
            else:
                for track in gpx.tracks:
                    for segment in track.segments:
                        for point in segment.points:
                            self.lats.append(point.latitude)
                            self.longs.append(point.longitude)
                            self.track.append((point.latitude, point.longitude))

    def parse_waypoints(self, waypointsfile):
        with open(waypointsfile, 'r') as gpx_file:
            gpx = gpxpy.parse(gpx_file)
            if not gpx.waypoints:
                raise ValueError("File doesn't contains any waypoints")
            else:
                for wp in gpx.waypoints:
                    self.waypoints.append([wp.name, wp.longitude, wp.latitude])

    def parse(self):
        current_terrain = self.waypoints[0][0]
        if current_terrain != Places.START.value:
            raise ValueError('No Start Error')
        else:
            buf = {}
            for point in self.waypoints[1:]:
                label = point[0]
                if label == Places.STOP.value:
                    break

                elif label in self.pcs:
                    ind = self.offset[label]
                    for alias, count in buf.items():
                        self.dict_ans[alias][ind] += count
                    buf = {}

                else:
                    if len(label) > 3:
                        animal_alias = label[0:3]
                        animal_count = label[3:]
                    else:
                        animal_alias = label[0:2]
                        animal_count = label[2:]
                    if animal_alias not in buf:
                        buf[animal_alias] = int(animal_count)
                    else:
                        buf[animal_alias] += int(animal_count)

    def prepare_context(self):
        context = {}
        for key, value in self.dict_ans.items():
            for i in range(3):
                context[f"{key}{i}"] = value[i]
        return context

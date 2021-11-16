import win32com.client
import time
import prometheus_client
import argparse
import yaml

import httpimport

with httpimport.github_repo('jewzaam', 'metrics-utility', 'utility', 'main'):
    import utility

# REFERENCE https://stackoverflow.com/questions/46163001/python-script-for-controlling-ascom-ccd-camera

REQUEST_TIME = prometheus_client.Summary('request_processing_seconds', 'Time spent processing request')

METRICS_FREQUENCY_SECONDS = 2

def getMetrics_Focuser(config):
    # https://ascom-standards.org/Help/Developer/html/T_ASCOM_DriverAccess_Focuser.htm

    # skip if there are no focusers configured
    if 'focuser' not in config:
        return

    for device in config['focuser']:
        # must have a driver for the device
        if 'driver' not in device:
            continue

        try:
            # if we cannot connect then we cannot collect metrics.
            focuser = win32com.client.Dispatch(device['driver'])
            focuser.Connected = True
        except Exception as e:
            print(e)
            return

        # if focuser is not connected bail
        if focuser.Connected == False:
            print(f"FAILURE: {device.driver} not connected")
            continue

        # collect all the data up front
        data = {}
        try:
            data["absolute"] = focuser.Absolute
        except Exception as e:
            print(e)

        try:
            data["is_moving"] = focuser.IsMoving
        except Exception as e:
            print(e)

        try:
            data["max_increment"] = focuser.MaxIncrement
        except Exception as e:
            print(e)

        try:
            data["max_step"] = focuser.MaxStep
        except Exception as e:
            print(e)

        try:
            data["name"] = focuser.Name
        except Exception as e:
            print(e)

        try:
            data["position"] = focuser.Position
        except Exception as e:
            print(e)

        """
        # Not Supported
        try:
            data["step_size"] = focuser.StepSize
        except Exception as e:
            print(e)
        """

        try:
            data["temp_comp"] = focuser.TempComp
        except Exception as e:
            print(e)

        try:
            data["temp_comp_available"] = focuser.TempCompAvailable
        except Exception as e:
            print(e)

        try:
            data["temperature"] = focuser.temperature
        except Exception as e:
            print(e)

        name=data['name']

        for key in ['max_increment', 'max_step', 'position', 'temperature']:
            if key in data:
                utility.set("ascom_focuser_data", data[key], {"name": name, "type": key})

        for key in ['absolute', 'is_moving', 'temp_comp', 'temp_comp_available']:
            if key in data and data[key]:
                utility.set("ascom_focuser_status", 1, {"name": name, "type": key})
            else:
                utility.set("ascom_focuser_status", 0, {"name": name, "type": key})

        # keep count of how many times we collect these metrics
        utility.inc("ascom_focuser_total", {"name": name})

def getMetrics_Switch(config):
    # https://ascom-standards.org/Help/Developer/html/T_ASCOM_DriverAccess_Switch.htm

    # skip if there are no switches configured
    if 'switch' not in config:
        return

    for device in config['switch']:
        # must have a driver for the device
        if 'driver' not in device:
            continue

        try:
            # if we cannot connect then we cannot collect metrics.
            switch = win32com.client.Dispatch(device['driver'])
            switch.Connected = True
        except Exception as e:
            print(e)
            return

        # if switch is not connected bail
        if switch.Connected == False:
            print(f"FAILURE: {device.driver} not connected")
            continue

        success = False
        for i in range(switch.MaxSwitch):
            try:
                # Getting value may fail, if so we don't care about that switch.
                try:
                    value = switch.GetSwitchValue(i)
                except:
                    continue
                name = switch.GetSwitchName(i)
                labelDict = {"device_name": switch.Description, "name": switch.Description, "index": i, "switch_name": name}
                utility.set("ascom_switch_data", value, labelDict)
                # if we set at least one metric we call it success
                success = True
            except Exception as e:
                print(e)
                pass
        
        if success:
            # keep count of how many times we collect these metrics
            utility.inc("ascom_switch_total", {"device_name": switch.Description, "name": switch.Description})

def getMetrics_Telescope(config):
    # https://ascom-standards.org/Help/Developer/html/T_ASCOM_DriverAccess_Telescope.htm

    # skip if there are no telescopes configured
    if 'telescope' not in config:
        return

    for device in config['telescope']:
        # must have a driver for the device
        if 'driver' not in device:
            continue

        try:
            # if we cannot connect then we cannot collect metrics.
            scope = win32com.client.Dispatch(device['driver'])
            scope.Connected = True
        except Exception as e:
            print(e)
            return

        # if scope is not connected bail
        if scope.Connected == False:
            print(f"FAILURE: {device.driver} not connected")
            continue
        
        # collect all the data up front
        data = {}
        try:
            data["alignment_mode"] = scope.AlignmentMode
        except Exception as e:
            print(e)

        try:
            data["altitude"] = scope.Altitude
        except Exception as e:
            print(e)

        try:
            data["at_home"] = scope.AtHome
        except Exception as e:
            print(e)

        try:
            data["at_park"] = scope.AtPark
        except Exception as e:
            print(e)

        try:
            data["azimuth"] = scope.Azimuth
        except Exception as e:
            print(e)

        try:
            data["declination"] = scope.Declination
        except Exception as e:
            print(e)

        try:
            data["declination_rate"] = scope.DeclinationRate
        except Exception as e:
            print(e)

        try:
            data["description"] = scope.Description
        except Exception as e:
            print(e)

        try:
            data["guide_rate_declination"] = scope.GuideRateDeclination
        except Exception as e:
            print(e)

        try:
            data["guide_rate_right_ascension"] = scope.GuideRateRightAscension
        except Exception as e:
            print(e)

        try:
            data["is_pulse_guiding"] = scope.IsPulseGuiding
        except Exception as e:
            print(e)

        try:
            data["name"] = scope.Name
        except Exception as e:
            print(e)

        try:
            data["right_ascension"] = scope.RightAscension
        except Exception as e:
            print(e)

        try:
            data["right_ascension_rate"] = scope.RightAscensionRate
        except Exception as e:
            print(e)

        try:
            data["side_of_pier"] = scope.SideOfPier
        except Exception as e:
            print(e)

        try:
            data["sidereal_time"] = scope.SiderealTime
        except Exception as e:
            print(e)

        try:
            data["site_elevation"] = scope.SiteElevation
        except Exception as e:
            print(e)

        try:
            data["site_latitude"] = scope.SiteLatitude
        except Exception as e:
            print(e)

        try:
            data["site_longitude"] = scope.SiteLongitude
        except Exception as e:
            print(e)

        try:
            data["slewing"] = scope.Slewing
        except Exception as e:
            print(e)

        try:
            data["slew_settle_time"] = scope.SlewSettleTime
        except Exception as e:
            print(e)

        try:
            data["tracking"] = scope.Tracking
        except Exception as e:
            print(e)

        try:
            data["tracking_rate"] = scope.TrackingRate
        except Exception as e:
            print(e)

        try:
            data["utc_date"] = scope.UTCDate
        except Exception as e:
            print(e)

        # these are not implemented:
        #'target_declination': scope.TargetDeclination,
        #'target_right_ascension': scope.TargetRightAscension,

        # proceed only if we have 'name'
        if 'name' in data:
            name=data['name']

            shared_labels={
                "name": name,
            }

            if "site_latitude" in data:
                shared_labels["latitude"] = data['site_latitude']
            if "site_longitude" in data:
                shared_labels["longitude"] = data['site_longitude']

            for key in ['slewing','at_home','at_park','alignment_mode','is_pulse_guiding','side_of_pier','slew_settle_time','tracking']:
                l = {"type": key}
                l.update(shared_labels)
                if key in data:
                    utility.set("ascom_telescope_status", data[key], l)
                else:
                    utility.set("ascom_telescope_status", None, l)

            for key in ['tracking_rate','declination_rate','right_ascension_rate','guide_rate_right_ascension','guide_rate_declination']:
                t = key.replace('_rate','')
                l = {"type": t}
                l.update(shared_labels)
                if key in data and data[key]:
                    utility.set("ascom_telescope_rate", data[key], l)
                # this caused failures, probably trying to delete a metric that didn't exist but that's a guess.
                #else:
                #    utility.set("ascom_telescope_rate", None, l)

            utility.set("ascom_telescope_site_elevation", data['site_elevation'], shared_labels)
            utility.set("ascom_telescope_site_latitude", data['site_latitude'], shared_labels)
            utility.set("ascom_telescope_site_longitude", data['site_longitude'], shared_labels)

            utility.set("ascom_telescope_altitude", data['altitude'], shared_labels)
            utility.set("ascom_telescope_azimuth", data['azimuth'], shared_labels)
            utility.set("ascom_telescope_declination", data['declination'], shared_labels)
            utility.set("ascom_telescope_right_ascension", data['right_ascension'], shared_labels)

            utility.inc("ascom_telescope_total", shared_labels)
        else:
            print("Telescope not connected, skipping.")

def getMetrics_Camera(config):
    # https://ascom-standards.org/Help/Developer/html/T_ASCOM_DriverAccess_Camera.htm


    # skip if there are no cameras configured
    if 'camera' not in config:
        return

    for device in config['camera']:
        # must have a driver for the device
        if 'driver' not in device:
            continue

        try:
            # if we cannot connect then we cannot collect metrics.
            camera = win32com.client.Dispatch(device['driver'])
            camera.Connected = True
        except Exception as e:
            print(e)
            return

        # if camera is not connected bail
        if camera.Connected == False:
            print(f"FAILURE: {device.driver} not connected")
            continue

        # collect all the data up front
        data = {}
        try:
            data["name"] = camera.Name
        except Exception as e:
            print(e)

        name=data['name']
        
        try:
            data["bin_x"] = camera.BinX
        except Exception as e:
            print(e)

        try:
            data["bin_Y"] = camera.BinY
        except Exception as e:
            print(e)

        try:
            data["camera_state"] = camera.CameraState
        except Exception as e:
            print(e)

        try:
            data["ccd_temperature"] = camera.CCDTemperature
        except Exception as e:
            print(e)

        try:
            data["can_abort_exposure"] = camera.CanAbortExposure
        except Exception as e:
            print(e)

        try:
            data["can_pulse_guide"] = camera.CanPulseGuide
        except Exception as e:
            print(e)

        try:
            data["can_set_ccd_temperature"] = camera.CanSetCCDTemperature
        except Exception as e:
            print(e)

        try:
            data["can_stop_exposure"] = camera.CanStopExposure
        except Exception as e:
            print(e)

        try:
            data["cooler_on"] = camera.CoolerOn
        except Exception as e:
            print(e)

        try:
            data["cooler_power"] = camera.CoolerPower
        except Exception as e:
            print(e)

        try:
            data["electrons_per_adu"] = camera.ElectronsPerADU
        except Exception as e:
            print(e)

        try:
            data["gain"] = camera.Gain
        except Exception as e:
            print(e)

        try:
            data["has_shutter"] = camera.HasShutter
        except Exception as e:
            print(e)

        try:
            data["is_pulse_guiding"] = camera.IsPulseGuiding
        except Exception as e:
            print(e)

        try:
            data["offset"] = camera.Offset
        except Exception as e:
            # fails on some cameras, just annoying in logs..
            pass

        for key in ['electrons_per_adu','gain','offset','bin_x','bin_Y','camera_state','cooler_power','ccd_temperature']:
            if key in data:
                utility.set("ascom_camera_data", data[key], {"name": name, "type": key})

        for key in ['cooler_on','can_abort_exposure','can_pulse_guide','can_set_ccd_temperature','can_stop_exposure','has_shutter','image_ready','is_pulse_guiding']:
            if key in data and data[key]:
                utility.set("ascom_camera_status", 1, {"name": name, "type": key})
            else:
                utility.set("ascom_camera_status", 0, {"name": name, "type": key})

        utility.inc("ascom_camera_total", {
            "name": name,
        })

@REQUEST_TIME.time()
def getMetrics(config):
    getMetrics_Focuser(config)
    getMetrics_Switch(config)
    getMetrics_Telescope(config)
    getMetrics_Camera(config)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Export logs as prometheus metrics.")
    parser.add_argument("--port", type=int, help="port to expose metrics on")
    parser.add_argument("--config", type=str, help="device configuration file")
    args = parser.parse_args()

    # Start up the server to expose the metrics.
    utility.metrics(args.port)

    # load configuration
    with open(args.config, 'r') as f:
        config = yaml.load(f, Loader=yaml.FullLoader)

    # Generate some requests.
    while True:
        getMetrics(config)
        time.sleep(METRICS_FREQUENCY_SECONDS)
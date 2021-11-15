import win32com.client
import argparse
import yaml

# Generate a configuration file based on CONNECTED devices.
# Scope is only those device types that are supported by this exporter.

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create ASCOM exporter config based on connected devices.")
    parser.add_argument("--config", type=str, help="output configuration file")
    args = parser.parse_args()
    
    configData = {}

    profile = win32com.client.Dispatch("ASCOM.Utilities.Profile")

    # find all the devices the exporter supports
    deviceTypes = ["telescope", "camera", "switch", "focuser"]
    for deviceType in deviceTypes:
        devices = profile.RegisteredDevices(deviceType)
        for device in devices:
            driver = device.Key
            name = device.Value
            if "Simulator" in driver or "Sim." in driver:
                # skip the simulators
                continue
            try:
                device = win32com.client.Dispatch(driver)
                # try connecting
                device.Connected = True
                if device.Connected == True:
                    print(f"Found {deviceType}: {driver}")
                    if deviceType not in configData:
                        configData[deviceType] = []
                    configData[deviceType].append({
                        'driver': str(driver), 
                        'name': str(name),
                    })
            except Exception as e:
                # ignore errors...
                #print(f"ERROR: {e}")
                pass

    # write config file
    with open(args.config, 'w') as out:
        yaml.dump(configData, out, default_flow_style=False)
# ascom-exporter

Export metrics from ASCOM devices in a form that Prometheus can scrape.

## Setup

1. install requirements
1. connect all devices, make sure they're setup in ASCOM
1. run configuration generator

```shell
# install required modules
pip3 install -r requirements.txt

# generate configuration
python ascom-config-generator.py --config config.yaml
```

## Usage

Simply run the exporter with a port and the config you've created / edited

```shell
python ascom-exporter.py --port 8001 --config config.yaml
```

## Verify

In your favorite browser look at the metrics endpoint.  If it's local, you can use http://localhost:8001

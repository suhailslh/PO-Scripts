# PO-Scripts

## Setup
In each folder (preferably in a virtual environment), run:
```python
py -m install -r requirements.txt
```

## To Read Communication Channels
add a `properties.json` file to the `Read-CC` folder like so...
```json
{
    "system/s" : "System1, System2",
    "protocol" : "http",
    "host/s"   : "<host1>:<port1>, <host2>:<port2>",
    "endpoint" : "/CommunicationChannelService/HTTPBasicAuth/",
    "dest"     : "path\\to\\destination\\file\\filename.xlsx"
}
```

## To Read Service Interfaces
1. add a `properties.json` file to the `Read-SI` folder similarly to how it's done for
   [Communication Channels](#to-read-communication-channels)
   except with the endpoint: `/rep/support/SimpleQuery`
3. add a `system-line-map.json` file to the same folder like so...
```json
{
    "SWC 1" : "System Line 1",
    "SWC 2" : "System Line 2",
    "SWC 3" : "System Line 3"
}
```

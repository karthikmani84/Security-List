import pandas
import json

excel_data_df = pandas.read_excel('Security-List.xls', sheet_name='Ingress')
excel_data_ds = pandas.read_excel('Security-List.xls', sheet_name='Egress')

#Ingress Rules JSON
json_str = excel_data_df.to_json(orient='records')

#Egress Rules JSON
json_sts = excel_data_ds.to_json(orient='records')

#Ingress Rule JSON Append

netData = json.loads(json_str)

ary = []
for x in netData:
    try:
        netJSON = {}
        netJSON["description"] = x["Description"]
        netJSON["isStateless"] = "false"
        netJSON["source-type"] = "CIDR_BLOCK"

        if (x["Protocol"] == "tcp"):
            #x["Protocol"] = "6"
            netJSON["icmp-options"] = None
            netJSON["udp-options"] = None
            netJSON["protocol"] = "6"
            netJSON["source"] = x["source"]
            netJSON["tcp-options"] = {"destinationPortRange":{"min":x["portl"],"max":x["portm"]}}
            

        elif (x["Protocol"] == "udp"):
            #x["Protocol"] = "17"
            netJSON["icmp-options"] = None
            netJSON["tcp-options"] = None
            netJSON["protocol"] = "17"
            netJSON["source"] = x["source"]
            netJSON["udp-options"] = {"destinationPortRange":{"min":x["portl"],"max":x["portm"]}}
            

        elif (x["Protocol"] == "icmp"):
            #x["Protocol"] = "1"
            netJSON["icmp-options"] = None
            netJSON["protocol"] = "1"
            netJSON["source"] = x["source"]
            netJSON["tcp-options"] = None
            netJSON["udp-options"] = None

        elif (x["Protocol"] == "all"):
            #x["Port"] = None
            netJSON["icmp-options"] = None
            netJSON["source"] = x["source"]
            netJSON["protocol"] = "all"
            netJSON["tcp-options"] = None
            netJSON["udp-options"] = None

        else:
            raise "Error processing the excel sheet."
    except:
        print ("Error processing the excel sheet.")

    ary.append(netJSON)

# Serializing json 
json_object = json.dumps(ary, indent = 4)
  
# Writing to sample.json
with open("ingress_list.json", "w") as outfile:
    outfile.write(json_object)

print("Ingress Rules JSON got created - Cheers!")

#Egress Rule JSON Append

Data = json.loads(json_sts)

ary = []
for x in Data:
    try:
        newJSON = {}
        newJSON["description"] = x["Description"]
        newJSON["isStateless"] = "false"
        newJSON["destination-type"] = "CIDR_BLOCK"

        if (x["Protocol"] == "tcp"):
            #x["Protocol"] = "6"
            newJSON["icmp-options"] = None
            newJSON["udp-options"] = None
            newJSON["protocol"] = "6"
            newJSON["destination"] = x["destination"]
            newJSON["tcp-options"] = {"destinationPortRange":{"min":x["portl"],"max":x["portm"]}}
            

        elif (x["Protocol"] == "udp"):
            #x["Protocol"] = "17"
            newJSON["icmp-options"] = None
            newJSON["tcp-options"] = None
            newJSON["protocol"] = "17"
            newJSON["destination"] = x["destination"]
            newJSON["udp-options"] = {"destinationPortRange":{"min":x["portl"],"max":x["portm"]}}
            

        elif (x["Protocol"] == "icmp"):
            #x["Protocol"] = "1"
            newJSON["icmp-options"] = None
            newJSON["protocol"] = "1"
            newJSON["destination"] = x["destination"]
            newJSON["tcp-options"] = None
            newJSON["udp-options"] = None

        elif (x["Protocol"] == "all"):
            #x["Port"] = None
            newJSON["icmp-options"] = None
            newJSON["destination"] = x["destination"]
            newJSON["protocol"] = "all"
            newJSON["tcp-options"] = None
            newJSON["udp-options"] = None

        else:
            raise "Error processing the excel sheet."
    except:
        print ("Error processing the excel sheet.")

    ary.append(newJSON)
#print (json.dumps(ary, indent=4))

# Serializing json 
newjson_object = json.dumps(ary, indent = 4)
  
# Writing to sample.json
with open("egress_list.json", "w") as outfile:
    outfile.write(newjson_object)

print("Egress Rules JSON got created in - Cheers!")
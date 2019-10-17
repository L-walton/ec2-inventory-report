#!/usr/bin/env python
"""
#!/usr/bin/env python
   Prerequisites:
       pip install openpyxl boto3

   Usage : python inventory.py AWS-profile-name

   This script create a XLSX report with EC2 instances details report on owned by AWS account.
"""
import awspricing
import boto3
import os
import re
import time
from datetime import datetime, timedelta
import logging
from botocore.exceptions import ClientError
import pprint
import tempfile
import yaml
import json
import csv
import sys
import string
from boto3.session import Session
from operator import itemgetter
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color

# static globals

os.environ['AWS_PROFILE'] = sys.argv[1]
os.environ['AWS_DEFAULT_REGION'] = "us-west-2"
timestr = time.strftime("%Y%m%d-%H%M%S")
logger = logging.getLogger()
logger.setLevel(logging.INFO)


def monitor_cw(instance_id, region):
    # returns cpu utilization from cloudwatch
    now = datetime.utcnow()
    past = now - timedelta(minutes=10090)
    future = now + timedelta(minutes=10)
    cwclient = boto3.client('cloudwatch', region_name=region)
    results = cwclient.get_metric_statistics(Namespace='AWS/EC2',
                                             MetricName='CPUUtilization',
                                             Dimensions=[{
                                                 'Name': 'InstanceId',
                                                 'Value': instance_id
                                             }],
                                             StartTime=past,
                                             EndTime=future,
                                             Period=86400,
                                             Statistics=['Average'])
    datapoints = results['Datapoints']
    load = ''
    if datapoints:
        last_datapoint = sorted(datapoints, key=itemgetter('Timestamp'))[-1]
        utilization = last_datapoint['Average']
        load = round((utilization), 2)
    return load


def monitor_ec2(region):
    # returns row with ec2 details.
    client = boto3.client('ec2', region_name=region)
    paginator = client.get_paginator('describe_instances')
    response_iterator = paginator.paginate()
    for page in response_iterator:
        for obj in page['Reservations']:
            for instance in obj['Instances']:
                InstanceName = None
                Platform = "linux"
                ID = instance['InstanceId']
                if instance['State']['Name'] != 'terminated':
                    for tag in instance["Tags"]:
                        try:
                            if tag["Key"] == 'Name':
                                InstanceName = tag["Value"]
                        except:
                            print "Tag Error", instance, tag
                #get attached volumes
                # ebsvolumes = list()
                # bDm = client.describe_instance_attribute(Attribute='blockDeviceMapping', InstanceId=ID)
                # for B in bDm['BlockDeviceMappings']:
                #     ebsvolumes.append(B['Ebs']['VolumeId'])
                PrivateIP = None
                PublicIPADDDR = None

                try:
                    ec2 = boto3.resource('ec2', region_name=region)
                except ClientError as ex:
                    print 'ec2'
                    error_message = ex.response['Error']['Message']
                    print 'setup_resource', error_message
                try:
                    InstanceDetails = ec2.Instance(instance['InstanceId'])
                except ClientError as ex:
                    print 'InstanceDetails'
                    error_message = ex.response['Error']['Message']
                    print 'Instance Details', error_message
                try:
                    ec2vol = list()
                    Volumes = InstanceDetails.volumes.all()
                    volume_ids = [v.id for v in Volumes]
                    for volume_id in volume_ids:
                        print volume_id
                        Vol = ec2.Volume(id=volume_id)
                        ec2vol.append(Vol.attachments[0][u'Device'])
                        ec2vol.append(Vol.size)
                except ClientError as ex:
                    print 'Volume'
                    error_message = ex.response['Error']['Message']
                    print 'Instance Volume Details', error_message

                print(InstanceName)

                try:
                    for inet in instance['NetworkInterfaces']:
                        if 'Association' in inet:
                            PublicIPADDDR = inet['Association']['PublicIp']

                        if 'Platform' in instance:
                            Platform = instance['Platform']

                        if 'PrivateIpAddress' in instance:
                            PrivateIP = instance['PrivateIpAddress']
                except:
                    error_message = ex.response['Error']['Message']
                    print 'Instance IP Details', error_message

                if PublicIPADDDR == None:
                    PublicIPADDDR = instance['PublicIpAddress']
                    PrivateIP = instance['PrivateIpAddress']

                try:

                    # print "Adding:", "Name:", InstanceName, "Region:", region, "Type:", instance["InstanceType"]
                    HW = data['compute']['models'][region][
                        instance["InstanceType"]]
                    #PRICE = data['compute']['prices'][region][instance["InstanceType"]]

                except:
                    print "Error in data", "Name:", InstanceName, "Region:", region, "Type:", instance[
                        "InstanceType"]

                ec2_offer = awspricing.offer('AmazonEC2')
                try:
                    on_demand_price = ec2_offer.ondemand_hourly(
                        instance["InstanceType"],
                        operating_system='Linux',
                        region=region)
                except:
                    print "on demand price error", instance[
                        "InstanceType"], region
                    on_demand_price = 0.000

                try:
                    reserved_price = ec2_offer.reserved_hourly(
                        instance["InstanceType"],
                        operating_system='Linux',
                        lease_contract_length='3yr',
                        offering_class='convertible',
                        purchase_option='Partial Upfront',
                        region=region)
                except:
                    print "reserved_price error", instance[
                        "InstanceType"], region
                    reserved_price = 0.000

                row = list()
                row.append(instance['Placement']['AvailabilityZone'])
                row.append(InstanceName)
                row.append(instance["InstanceId"])
                row.append(instance["InstanceType"])
                row.append(Platform)
                row.append(PublicIPADDDR)
                row.append(PrivateIP)
                row.append(instance['State']['Name'])
                row.append(instance['LaunchTime'])
                row.append(Account)
                row.append(HW['CPU'])
                row.append(monitor_cw(instance["InstanceId"], region))
                row.append(HW['ECU'])
                row.append(HW['memoryGiB'])
                row.append(on_demand_price)
                row.append(round(reserved_price, 4))
                row.extend(ec2vol)
                ws.append(row)

def get_regions():
    """
        List all AWS region available
        :return: list of regions
    """
    client = boto3.client('ec2')
    regions = [
        region['RegionName'] for region in client.describe_regions()['Regions']
    ]
    return regions


def format_xlsx(ws):
    #
    #       Formating XLSX
    #   TOP Row
    try:
        for rows in ws.iter_rows(min_row=1, max_row=1, min_col=1):
            for cell in rows:
                cell.fill = PatternFill("solid", fgColor="0066a1")
                cell.font = Font(color="00FFFFFF", bold=True)
#   Format entire table
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
    except:
        print "rows", rows

    # set wrapText
    try:
        alignment = Alignment(wrap_text=True)
        for col in ws.iter_cols(min_row=1, min_col=1, max_col=26):
            for cell in col:
                cell.alignment = alignment
    except:
        print "col", col


# set weight
    for col in ws.iter_cols(min_row=1, min_col=1, max_col=26):
        max_length = 0
        c = col[0].column
        column = get_column_letter(c)
        for cell in col[1:]:
            cell.border = thin_border
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
                    if ' ' in str(cell.value):
                        max_length = ((len(cell.value) + 2) * 1.4) / 2
            except:
                pass
        adjusted_width = (max_length + 3)
        ws.column_dimensions[column].width = adjusted_width


def init_moniroting():
    """
        Script initialisation
        :retur: None
    """
    # Setting regions
    global region_list
    region_list = get_regions()


if __name__ == '__main__':
    import sys
    logging.basicConfig(level=logging.WARNING)
    try:
        print('Loading price.json')
        with open('price.json') as json_file:
            data = json.load(json_file)
        wb = Workbook()
        # grab the active worksheet
        ws = wb.active
        header = [
            'Placement', 'Name', 'Instance ID', 'Instance Type', 'Platform',
            'Public IP', 'Private IP', 'Instance State', 'LaunchTime',
            'AWS Account', 'CPU', 'CPU Utilization Avg', 'ECU', 'memory GiB',
            'On demand price', 'Reserved price', 'Volume', 'Size GiB',
            'Volume', 'Size GiB', 'Volume', 'Size GiB', 'Volume', 'Size GiB',
            'Volume', 'Size GiB'
        ]
        ws.append(header)
        global Account
        iam = boto3.client("iam")
        paginator = iam.get_paginator('list_account_aliases')
        for response in paginator.paginate():
            Account = "\n".join(response['AccountAliases'])
            # Calling functions for every region
        init_moniroting()
        for region in region_list:
            monitor_ec2(region)
        ws = format_xlsx(ws)
        wb.save("inventory-" + Account + "-" + timestr + ".xlsx")
    except ClientError as e:
        logger.error(e)
    except Exception as err:
        logger.error(err)

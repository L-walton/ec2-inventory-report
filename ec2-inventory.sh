#! /usr/local/bin/bash

export AWSPRICING_USE_CACHE="1"
export AWSPRICING_CACHE_PATH="./"
export AWSPRICING_CACHE_MINUTES="262800"

./ec2-inventory.py default

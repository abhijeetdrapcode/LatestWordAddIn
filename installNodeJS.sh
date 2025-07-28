#!/bin/bash

# Check if Node.js is installed
if command -v node > /dev/null 2>&1; then
    echo "Node.js is already installed in the system."
else
    echo "Node.js is not installed. Installing now..."
    
    # Update package list
    sudo apt update
    
    # Install Node.js
    sudo apt install -y nodejs
    
    # Install npm (Node Package Manager)
    sudo apt install -y npm
    
    # Verify installation
    if command -v node > /dev/null 2>&1; then
        echo "Node.js has been successfully installed!"
    else
        echo "Failed to install Node.js."
    fi
fi

#!/bin/bash

# Check if package.json exists (to confirm we're in a Node.js project directory)
if [ -f "package.json" ]; then
    echo "Found package.json, proceeding with npm install..."
    
    # Run npm install to install dependencies
    npm install
    
    # Check if npm install was successful
    if [ $? -eq 0 ]; then
        echo "npm install completed successfully."
        
        # Run npm start to start the application
        echo "Running npm start..."
        npm start
    else
        echo "npm install failed. Please check for errors."
    fi
else
    echo "package.json not found. Please make sure you are in the root directory of a Node.js project."
fi

#!/bin/bash
# Script to fix the monthly dashboard issue

# Set the base directory
BASE_DIR="/Users/mpriessner/windsurf_repos/dictate_tool"
REACT_DIR="$BASE_DIR/dashboard-react"

# Step 1: Build the React app
echo "Building React app..."
cd "$REACT_DIR"
npm run build

# Step 2: Get the current JS filename from the build output
JS_FILE=$(find "$REACT_DIR/build/static/js" -name "main.*.js" | grep -v ".map" | grep -v "LICENSE" | head -1)
CSS_FILE=$(find "$REACT_DIR/build/static/css" -name "main.*.css" | grep -v ".map" | head -1)

JS_FILENAME=$(basename "$JS_FILE")
CSS_FILENAME=$(basename "$CSS_FILE")

echo "Found JS file: $JS_FILENAME"
echo "Found CSS file: $CSS_FILENAME"

# Step 3: Update the dashboard.html file with the correct filenames
echo "Updating dashboard.html..."
sed -i.bak "s|static/js/main.*.js|static/js/$JS_FILENAME|" "$BASE_DIR/dashboard.html"
sed -i.bak "s|static/css/main.*.css|static/css/$CSS_FILENAME|" "$BASE_DIR/dashboard.html"

# Step 4: Create the static directories if they don't exist
echo "Creating static directories..."
mkdir -p "$BASE_DIR/static/js" "$BASE_DIR/static/css"

# Step 5: Copy the built files to the static directories
echo "Copying built files to static directories..."
cp "$REACT_DIR/build/static/js/$JS_FILENAME" "$BASE_DIR/static/js/"
cp "$REACT_DIR/build/static/css/$CSS_FILENAME" "$BASE_DIR/static/css/"

# Copy map and license files if they exist
[ -f "$REACT_DIR/build/static/js/$JS_FILENAME.map" ] && cp "$REACT_DIR/build/static/js/$JS_FILENAME.map" "$BASE_DIR/static/js/"
[ -f "$REACT_DIR/build/static/js/$JS_FILENAME.LICENSE.txt" ] && cp "$REACT_DIR/build/static/js/$JS_FILENAME.LICENSE.txt" "$BASE_DIR/static/js/"
[ -f "$REACT_DIR/build/static/css/$CSS_FILENAME.map" ] && cp "$REACT_DIR/build/static/css/$CSS_FILENAME.map" "$BASE_DIR/static/css/"

echo "Monthly dashboard fix complete!"
echo "Please restart the app and try viewing the monthly dashboard again."
#!/bin/bash
# Generate a QR code PNG on the Desktop for the given URL.
# Usage: ./generate-qr.sh "https://example.com"

set -e

URL="$1"
if [ -z "$URL" ]; then
  echo "Usage: $0 <url>"
  exit 1
fi

OUTPUT="$HOME/Desktop/jungle-safari-qr.png"

qrencode -o "$OUTPUT" -s 12 -m 2 -l H "$URL"

echo "QR code saved to: $OUTPUT"
echo "URL encoded: $URL"

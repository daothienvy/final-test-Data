#!/bin/bash
# Start auto-reconnecting tunnel for Dashboard
OUTPUT_FILE="/Users/ThienTuCorp/Downloads/Final test Data/public_url.txt"
echo "Đang khởi tạo liên kết công khai..." > "$OUTPUT_FILE"

while true; do
    ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=30 -o ServerAliveCountMax=3 -R 80:localhost:8080 nokey@localhost.run 2>&1 | while read -r line; do
        echo "$line"
        if echo "$line" | grep -q "https://.*\.lhr\.life"; then
            URL=$(echo "$line" | grep -o 'https://[^ ]*\.lhr\.life')
            echo "$URL" > "$OUTPUT_FILE"
            echo "PUBLIC_URL_READY: $URL"
        fi
    done
    sleep 3
done

#!/bin/bash
# Azure App Service (Linux, Node) does not ship the shared libraries that
# @sparticuz/chromium's bundled Chromium binary needs to run — that package
# targets AWS Lambda, where these are preinstalled. Install them here, once,
# before every container start (nothing outside /home persists between
# restarts, so this has to run on boot rather than at build time).
apt-get update -qq > /dev/null 2>&1
apt-get install -y --no-install-recommends \
  libnspr4 libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 \
  libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 libxrandr2 \
  libgbm1 libasound2 libpango-1.0-0 libcairo2 libxext6 libx11-6 \
  libxcb1 > /dev/null 2>&1

exec node server.js

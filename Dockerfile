# syntax=docker/dockerfile:1
FROM node:20-alpine AS base

WORKDIR /app

# Install production dependencies only
COPY package*.json ./
RUN npm ci --omit=dev

# Copy source
COPY src/ ./src/

# Expose bot port
EXPOSE 3978

# Health-check endpoint
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget -qO- http://localhost:3978/health || exit 1

CMD ["node", "src/index.js"]

FROM node:20-alpine AS builder

WORKDIR /app

# Copy dependency manifests
COPY package*.json ./

# Install dependencies
RUN npm ci

# Copy application source code
COPY . .

# Build TypeScript code
RUN npm run build

# --- Production Image ---
FROM node:20-alpine AS runner

WORKDIR /app

COPY package*.json ./

# Install only production dependencies
RUN npm ci --only=production

# Copy compiled TypeScript code from builder stage
COPY --from=builder /app/dist ./dist

# Create empty keys directory (will be overridden by volume mount)
RUN mkdir -p /app/keys

CMD ["node", "dist/index.js"]
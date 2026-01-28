# Build stage
FROM node:20-alpine AS builder

WORKDIR /app

# Build arguments (set these in Easypanel)
ARG AZURE_OPENAI_ENDPOINT
ARG AZURE_OPENAI_API_KEY
ARG SEARXNG_PROXY_URL
ARG AVAILABLE_MODELS
ARG DEFAULT_MODEL_ID

# Copy package files
COPY package.json ./

# Install dependencies
RUN npm install

# Copy source files
COPY . .

# Create .env file from build args
RUN echo "AZURE_OPENAI_ENDPOINT=${AZURE_OPENAI_ENDPOINT}" > .env && \
    echo "AZURE_OPENAI_API_KEY=${AZURE_OPENAI_API_KEY}" >> .env && \
    echo "SEARXNG_PROXY_URL=${SEARXNG_PROXY_URL}" >> .env && \
    echo "AVAILABLE_MODELS=${AVAILABLE_MODELS}" >> .env && \
    echo "DEFAULT_MODEL_ID=${DEFAULT_MODEL_ID}" >> .env

# Build the application
RUN npm run build

# Production stage
FROM nginx:alpine

# Copy custom nginx config
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copy built files from builder stage
COPY --from=builder /app/dist /usr/share/nginx/html

# Expose port 80 (Easypanel handles SSL)
EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]

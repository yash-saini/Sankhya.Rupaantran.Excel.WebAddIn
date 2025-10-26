# Stage 1: Build the Blazor app
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy the project file and restore dependencies
COPY *.csproj .
RUN dotnet restore

# Copy the rest of the source code and publish
COPY . .
RUN dotnet publish -c Release -o /app/publish

# Stage 2: Create the final Nginx image to serve the app
FROM nginx:alpine AS final
WORKDIR /usr/share/nginx/html

# Copy the published Blazor app files from the build stage
COPY --from=build /app/publish/wwwroot .

# Copy a custom Nginx configuration file (we'll create this next)
COPY nginx.conf /etc/nginx/nginx.conf

# Expose the port Nginx runs on
EXPOSE 80
FROM mcr.microsoft.com/dotnet/sdk:5.0 AS build-env
WORKDIR /app

# Copy everything else and build
COPY . .
RUN dotnet publish WebScrapingCar/ -c Release -o out --disable-parallel

# Build runtime image
FROM mcr.microsoft.com/dotnet/aspnet:5.0 AS base
WORKDIR /app
COPY --from=build-env /app/out .
ENTRYPOINT ["dotnet", "WebScrapingCar.dll"]
# CMD ASPNETCORE_URLS=http://*:$PORT dotnet WebScrapingCar.dll

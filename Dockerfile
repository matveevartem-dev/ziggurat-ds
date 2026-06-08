FROM mcr.microsoft.com/dotnet/sdk:10.0-alpine AS build
WORKDIR /src
COPY DocumentService.csproj .
RUN dotnet restore
COPY . .
RUN dotnet publish -c release -o /ds

FROM mcr.microsoft.com/dotnet/aspnet:10.0
WORKDIR /ds
COPY --from=build /ds .
ENTRYPOINT ["dotnet", "DocumentService.dll"]
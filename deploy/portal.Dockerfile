FROM mcr.microsoft.com/dotnet/sdk:8.0

WORKDIR /app

COPY ITSupportSystem/*.csproj ITSupportSystem/
RUN dotnet restore ITSupportSystem/ITSupportSystem.csproj

COPY ITSupportSystem/ ITSupportSystem/
RUN dotnet publish ITSupportSystem/ITSupportSystem.csproj -c Release -o /app/ITSupportSystem/publish

RUN apt-get update \
    && apt-get install -y curl ca-certificates gnupg \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

COPY Website/package*.json Website/
RUN npm --prefix Website install --omit=dev

COPY Website/ Website/
COPY scripts/start-portal-and-monitor.sh /app/start.sh
RUN chmod +x /app/start.sh

ENV PORT=8080
ENV SUPPORT_SYSTEM_DLL_PATH=/app/ITSupportSystem/publish/ITSupportSystem.dll
ENV SUPPORT_SETTINGS_PATH=/app/ITSupportSystem/appsettings.json

EXPOSE 8080

CMD ["/app/start.sh"]

FROM mcr.microsoft.com/vscode/devcontainers/java:latest

SHELL ["/bin/bash", "-c"]
RUN . ${SDKMAN_DIR}/bin/sdkman-init.sh && sdk install scala
RUN . ${SDKMAN_DIR}/bin/sdkman-init.sh && sdk install gradle



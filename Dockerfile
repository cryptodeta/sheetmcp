# syntax=docker/dockerfile:1
FROM python:3.12-slim AS build
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv
WORKDIR /src
COPY . .
RUN uv build --wheel --out-dir /wheels

FROM python:3.12-slim
LABEL org.opencontainers.image.title="sheetmcp" \
      org.opencontainers.image.description="Excel/CSV MCP server: formulas, SQL across files, safe writes" \
      org.opencontainers.image.licenses="MIT"
COPY --from=build /wheels /wheels
RUN pip install --no-cache-dir /wheels/sheetmcp-*.whl \
    && rm -rf /wheels \
    && useradd --create-home --uid 1000 app && mkdir -p /data && chown app /data
USER app
ENV SHEETMCP_ROOTS=/data
VOLUME ["/data"]
EXPOSE 8000
ENTRYPOINT ["sheetmcp"]
# default: stdio. For HTTP: docker run -p 8000:8000 -v $PWD:/data sheetmcp --transport http --host 0.0.0.0
CMD []

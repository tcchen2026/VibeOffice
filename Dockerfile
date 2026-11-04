# VibeOffice.work as a static site served by nginx (docs/suite.md, Self-hosting with Docker):
#   docker build -t vibeoffice .
#   docker run -d --name vibeoffice -p 8080:80 vibeoffice      then open http://localhost:8080/
# Serve it to other machines over https: on plain http the browser keeps some features off (docs/suite.md, Plain http).
FROM docker.io/library/nginx:stable-alpine
COPY tools/docker/nginx.conf /etc/nginx/conf.d/default.conf
COPY public/ /usr/share/nginx/html/

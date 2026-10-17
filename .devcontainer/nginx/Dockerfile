FROM nginx:alpine

# Install OpenSSL for certificate generation
RUN apk add --no-cache openssl

# Copy rebuild-certs.sh script to the container
COPY nginx/rebuild-certs.sh /usr/local/bin/rebuild-certs.sh
RUN chmod +x /usr/local/bin/rebuild-certs.sh

# Generate certificates at build time
ARG SSL_DAYS=365
ARG SSL_SUBJECT_BASE="/C=JP/ST=Origami/L=Layer/O=ValleyMountain/OU=StackGeometry"
RUN SSL_DAYS=${SSL_DAYS} SSL_SUBJECT_BASE=${SSL_SUBJECT_BASE} /usr/local/bin/rebuild-certs.sh

EXPOSE 80 443

CMD ["nginx", "-g", "daemon off;"]

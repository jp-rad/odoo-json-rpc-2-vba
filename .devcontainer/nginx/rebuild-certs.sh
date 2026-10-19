#!/bin/sh
set -e

# =====================================
# Configuration
# =====================================
SSL_DAYS=${SSL_DAYS:-365}

# Unified subject base
SSL_SUBJECT_BASE=${SSL_SUBJECT_BASE:-"/C=JP/ST=Origami/L=Layer/O=ValleyMountain/OU=StackGeometry"}

# =====================================
# Directory to store SSL certificates
# =====================================
SSL_DIR="/etc/nginx/ssl"
echo "Rebuilding certificates in ${SSL_DIR} ..."
mkdir -p "${SSL_DIR}"

# =====================================
# Generate Root CA
# =====================================
SSL_SUBJECT="${SSL_SUBJECT_BASE}/CN=Development-Root-CA"
openssl req -x509 -new -nodes -days "${SSL_DAYS}" \
  -keyout "${SSL_DIR}/ca.key" \
  -out "${SSL_DIR}/ca.crt" \
  -subj "${SSL_SUBJECT}"

# =====================================
# Generate server key and CSR
# =====================================
SSL_SUBJECT="${SSL_SUBJECT_BASE}/CN=localhost-Developer"
openssl req -new -nodes \
  -keyout "${SSL_DIR}/server.key" \
  -out "${SSL_DIR}/server.csr" \
  -subj "${SSL_SUBJECT}"

# =====================================
# SAN configuration
# =====================================
cat <<EOF > "${SSL_DIR}/san.ext"
authorityKeyIdentifier=keyid,issuer
basicConstraints=CA:FALSE
keyUsage = digitalSignature, nonRepudiation, keyEncipherment, dataEncipherment
subjectAltName = @alt_names

[alt_names]
DNS.1 = localhost
IP.1  = 127.0.0.1
IP.2  = ::1
EOF

# =====================================
# Sign server certificate with Root CA
# =====================================
openssl x509 -req -in "${SSL_DIR}/server.csr" \
  -CA "${SSL_DIR}/ca.crt" \
  -CAkey "${SSL_DIR}/ca.key" \
  -CAcreateserial \
  -out "${SSL_DIR}/server.crt" \
  -days "${SSL_DAYS}" -sha256 \
  -extfile "${SSL_DIR}/san.ext"

echo "Certificate regeneration completed."

# =====================================
# Reload nginx if running
# =====================================
if pidof nginx >/dev/null 2>&1; then
  echo "Reloading nginx..."
  nginx -s reload || true
else
  echo "nginx is not running; reload skipped."
fi

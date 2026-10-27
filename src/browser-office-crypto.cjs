'use strict';

// Browser-only wiring for the existing Office AES/hash implementations.
// Node CLI/server/desktop continue to use Node's native crypto module.
const aes = require('browserify-aes/browser');

exports.createHash = require('create-hash');
exports.createHmac = require('create-hmac');
exports.createCipheriv = aes.createCipheriv;
exports.createDecipheriv = aes.createDecipheriv;
exports.randomBytes = require('randombytes');
exports.getHashes = () => ['sha1', 'sha224', 'sha256', 'sha384', 'sha512', 'md5', 'rmd160'];

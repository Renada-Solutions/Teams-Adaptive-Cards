const { BlobServiceClient } = require('@azure/storage-blob');

const CONTAINER = 'botstate';
const BLOB = 'conversationReference.json';

let cached = null;

function client() {
    return BlobServiceClient
        .fromConnectionString(process.env.AzureWebJobsStorage)
        .getContainerClient(CONTAINER)
        .getBlockBlobClient(BLOB);
}

async function load() {
    if (cached) return cached;
    try {
        const buf = await client().downloadToBuffer();
        cached = JSON.parse(buf.toString('utf8'));
        return cached;
    } catch (e) {
        if (e.statusCode === 404 || e.code === 'BlobNotFound' || e.code === 'ContainerNotFound') {
            return null;
        }
        throw e;
    }
}

async function save(ref) {
    const svc = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
    await svc.getContainerClient(CONTAINER).createIfNotExists();
    const data = JSON.stringify(ref);
    await client().upload(data, Buffer.byteLength(data, 'utf8'), {
        blobHTTPHeaders: { blobContentType: 'application/json' }
    });
    cached = ref;
}

module.exports = { load, save };

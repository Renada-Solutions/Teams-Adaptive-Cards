const { BlobServiceClient } = require('@azure/storage-blob');

// One blob per chat: where its card was posted, plus the original payload so the card can be rebuilt
const CONTAINER = 'botstate';

let containerReady = false;

function container() {
    return BlobServiceClient
        .fromConnectionString(process.env.AzureWebJobsStorage)
        .getContainerClient(CONTAINER);
}

function blob(chatId) {
    const safeId = String(chatId).replace(/[^0-9A-Za-z_-]/g, '');
    return container().getBlockBlobClient(`chats/${safeId}.json`);
}

async function streamToString(stream) {
    const chunks = [];
    for await (const chunk of stream) chunks.push(Buffer.from(chunk));
    return Buffer.concat(chunks).toString('utf8');
}

// Returns { record, etag }, or null if no card has been sent for this chat
async function getChat(chatId) {
    try {
        const response = await blob(chatId).download();
        const record = JSON.parse(await streamToString(response.readableStreamBody));
        return { record, etag: response.etag };
    } catch (e) {
        if (e.statusCode === 404 || e.code === 'BlobNotFound' || e.code === 'ContainerNotFound') {
            return null;
        }
        throw e;
    }
}

// Save a chat record. With etag, only succeeds if nobody else has changed it since it was read;
// returns false in that case so the caller can skip a duplicate update.
async function saveChat(chatId, record, etag) {
    if (!containerReady) {
        await container().createIfNotExists();
        containerReady = true;
    }
    const data = JSON.stringify(record);
    try {
        await blob(chatId).upload(data, Buffer.byteLength(data, 'utf8'), {
            blobHTTPHeaders: { blobContentType: 'application/json' },
            conditions: etag ? { ifMatch: etag } : undefined
        });
        return true;
    } catch (e) {
        if (etag && (e.statusCode === 412 || e.code === 'ConditionNotMet')) return false;
        throw e;
    }
}

module.exports = { getChat, saveChat };

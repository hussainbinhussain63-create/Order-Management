import mongoose from 'mongoose';

const MONGODB_URI = process.env.MONGODB_URI;

if (!MONGODB_URI) {
  throw new Error('Please define the MONGODB_URI environment variable inside .env');
}

function formatMongoUri(uri: string): string {
  if (!uri) return uri;
  try {
    const protocolMatch = uri.match(/^(mongodb(?:\+srv)?:\/\/)/);
    if (!protocolMatch) return uri;
    const protocol = protocolMatch[1];
    const withoutProtocol = uri.slice(protocol.length);

    const slashIdx = withoutProtocol.indexOf('/');
    const qIdx = withoutProtocol.indexOf('?');
    let endOfAuth = withoutProtocol.length;
    if (slashIdx !== -1) endOfAuth = Math.min(endOfAuth, slashIdx);
    if (qIdx !== -1) endOfAuth = Math.min(endOfAuth, qIdx);

    const authority = withoutProtocol.slice(0, endOfAuth);
    const pathAndQuery = withoutProtocol.slice(endOfAuth);

    const lastAtIdx = authority.lastIndexOf('@');
    if (lastAtIdx === -1) return uri;

    const userInfo = authority.slice(0, lastAtIdx);
    const host = authority.slice(lastAtIdx + 1);

    const firstColonIdx = userInfo.indexOf(':');
    if (firstColonIdx === -1) {
      return `${protocol}${encodeURIComponent(decodeURIComponent(userInfo))}@${host}${pathAndQuery}`;
    }

    const username = userInfo.slice(0, firstColonIdx);
    const password = userInfo.slice(firstColonIdx + 1);

    const safeUsername = encodeURIComponent(decodeURIComponent(username));
    const safePassword = encodeURIComponent(decodeURIComponent(password));

    return `${protocol}${safeUsername}:${safePassword}@${host}${pathAndQuery}`;
  } catch (e) {
    console.error("URI formatting failed, using raw URI:", e);
    return uri;
  }
}

/**
 * Global is used here to maintain a cached connection across hot reloads
 * in development. This prevents connections growing exponentially
 * during API Route usage.
 */
interface MongooseGlobal {
  conn: typeof mongoose | null;
  promise: Promise<typeof mongoose> | null;
}

declare global {
  // eslint-disable-next-line no-var
  var mongooseGlobal: MongooseGlobal | undefined;
}

let cached = global.mongooseGlobal;

if (!cached) {
  cached = { conn: null, promise: null };
  global.mongooseGlobal = cached;
}

const mongooseCached = cached as MongooseGlobal;

export async function dbConnect() {
  if (mongooseCached.conn) {
    return mongooseCached.conn;
  }

  if (!mongooseCached.promise) {
    const opts = {
      bufferCommands: false,
    };

    const formattedUri = formatMongoUri(MONGODB_URI!);
    const maskedUri = formattedUri.replace(/\/\/([^:]+):([^@]+)@/, '//***:***@');
    console.log('[MongoDB] Creating new connection promise to:', maskedUri);

    mongooseCached.promise = mongoose.connect(formattedUri, opts).then((m) => {
      console.log('[MongoDB] Mongoose successfully connected.');
      return m;
    }).catch((err) => {
      console.error('[MongoDB] Mongoose connection promise rejected:', err);
      throw err;
    });
  }

  try {
    mongooseCached.conn = await mongooseCached.promise;
    console.log('[MongoDB] Active connection established.');
  } catch (e) {
    console.error('[MongoDB] Error resolving mongoose connection promise:', e);
    mongooseCached.promise = null;
    throw e;
  }

  return mongooseCached.conn;
}

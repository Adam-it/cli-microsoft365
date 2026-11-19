// only needed for instanceof checks and as base classes

export class Stream { }
export class Readable extends Stream { }
export class Writable extends Stream { }
export class Duplex extends Stream { }
export class Transform extends Stream { }
export class PassThrough extends Stream { }

export default { Stream, Readable, Writable, Duplex, Transform, PassThrough };

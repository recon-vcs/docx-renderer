export enum Overflow {
	NONE = 'none',         // confirmed no overflow
	SELF = 'self',         // this element caused overflow
	PARTIAL = 'partial',   // some children overflowed
	FULL = 'full',         // all children overflowed
	UNCHECKED = 'unchecked', // no measurement taken
	SKIP = 'skip',         // skip this element during overflow tracking
}

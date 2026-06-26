import * as _ from 'lodash-es';

export class Length {
	constructor(readonly value: number, readonly type?: string) {
	}

	static parse(text: string): Length {
		const value = parseFloat(text);
		const type = /p[tx]$/i.exec(text)?.[0];
		return new Length(value, type);
	}

	static from(val: any): Length {
		if (_.isString(val)) {
			return Length.parse(val);
		}
		if (val instanceof Length) {
			return val;
		}

		return null;
	}

	// add
	add(length: Length): Length {
		if (length.type !== this.type) {
			throw new Error("Can't do math on different types");
		}

		return new Length(this.value + length.value, this.type);
	}

	// minus
	minus(length: Length): Length {
		if (length.type !== this.type) {
			throw new Error("Can't do math on different types");
		}

		return new Length(this.value - length.value, this.type);
	}

	// multiply
	multiply(val: number): Length {
		return new Length(this.value * val, this.type);
	}

	// divide
	divide(val: number): Length {
		return new Length(this.value / val, this.type);
	}

	valueOf() {
		return this.value;
	}

	toString(): string {
		return `${this.value.toFixed(2)}${this.type ?? ''}`;
	}
}

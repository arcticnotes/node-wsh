export class Collection {

	#ticket;
	#type;

	constructor( ticket, type) {
		this.#ticket = ticket;
		this.#type = type;
	}

	get length() {
		return this.#ticket.get( 'Count');
	}

	get( indexBase0) {
		return this.#ticket.get( 'Item', t2 => t2.apply.bind( t2))( this, [ indexBase0 + 1], t3 => new this.#type( t3));
	}

	*[ Symbol.iterator]() {
		const length = this.length; // expect anomoly if the underlying collect is changing
		for( let i = 0; i < length; i++)
			yield this.get( i);
	}
}

export interface Deck {
	category: string;
	/** Every speaking entry that linked to this deck, newest first. */
	events: DeckEvent[];
	/** A stable, readable id: the newest event's date and the talk title as a URL slug. */
	id: string;
	/** The deck's 1-based position, newest first: the numbering of earlier ad hoc downloads. */
	number: number;
	/** SHA-256 of the downloaded .pptx, once it's in the corpus. */
	sha256?: string;
	/** How many slides the downloaded .pptx has, once it's in the corpus. */
	slides?: number;
	title: string;
	/** The deck's OneDrive or SharePoint share link. */
	url: string;
}

export interface DeckEvent {
	date: string;
	event: string;
}

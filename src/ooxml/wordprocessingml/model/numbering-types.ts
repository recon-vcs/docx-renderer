export interface NumberingPicBullet {
	id: number;
	src: string;
	style?: string;
}

export interface IDomNumbering {
	id: string;
	level: number;
	start: number;
	pStyleName: string;
	pStyle: Record<string, string>;
	rStyle: Record<string, string>;
	levelText?: string;
	suff: string;
	format?: string;
	bullet?: NumberingPicBullet;
}

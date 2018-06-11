/*--------------------------------------------------------------------------
   Definitions of the supported orderBy parameters.
 -------------------------------------------------------------------------*/

export enum Parameter {
	lastModifiedTime,
	createdTime,
	name
}

export enum Direction {
	asc,
	desc
}

export class OrderBy {
	public parameter: Parameter;
	public direction?: Direction;
}

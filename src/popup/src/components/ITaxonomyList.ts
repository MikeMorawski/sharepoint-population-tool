import { ITermInfo } from "@pnp/sp/taxonomy";

export interface ITaxonomyList {
    [id: string]: ITermInfo[];
  }
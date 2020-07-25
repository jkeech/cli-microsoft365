// TODO: after removing Vorpal, rename to CommandOptions
export interface CliCommandOption {
  autocomplete: string[] | undefined;
  long?: string;
  name: string;
  required: boolean;
  short?: string;
}
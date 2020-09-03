export interface CommandOption {
  autocomplete: string[] | undefined;
  long?: string;
  name: string;
  required: boolean;
  short?: string;
}
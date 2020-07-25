import Command from "../Command";
import { CliCommandOption } from "./CommandOption";

// TODO: after removing Vorpal, rename to CommandInfo
export interface CliCommandInfo {
  aliases: string[] | undefined;
  command: Command;
  name: string;
  options: CliCommandOption[];
}
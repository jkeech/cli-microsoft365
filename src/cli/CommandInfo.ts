import Command from "../Command";
import { CommandOption } from "./CommandOption";

export interface CommandInfo {
  aliases: string[] | undefined;
  command: Command;
  name: string;
  options: CommandOption[];
}
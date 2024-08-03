/* eslint-disable @typescript-eslint/no-explicit-any */
import BasicSearch from "./examples/BasicSearch"
import BasicUserSearch from "./examples/BasicUserSearch"
import CurrentUserPersona from "./examples/CurrentUserPersona"
import ErrorHandling from "./examples/ErrorHandling"
import GroupsAndUsers from "./examples/GroupsAndUsers"
import JsonFileData from "./examples/JsonFileData"
import ListItems from "./examples/ListItems"
import Navigation from "./examples/Navigation"
import OptionProvider from "./examples/OptionProvider"
import PagedListItems from "./examples/PagedListItems"
import UserRoles from "./examples/UserRoles"

const EXAMPLES: [string, any][] = [
  ["Basic search", BasicSearch],
  ["Basic user search", BasicUserSearch],
  ["Current user persona", CurrentUserPersona],
  ["Error handling", ErrorHandling],
  ["Groups and users", GroupsAndUsers],
  ["Json file viewer", JsonFileData],
  ["List items", ListItems],
  ["Navigation", Navigation],
  ["Option provider", OptionProvider],
  ["Paged list items", PagedListItems],
  ["User roles", UserRoles],
];

export default EXAMPLES;


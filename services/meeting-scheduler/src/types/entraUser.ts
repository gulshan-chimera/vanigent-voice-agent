/** Directory candidate only; an email identity does not prove mailbox availability. */
export interface EntraUser {
  id: string;
  displayName: string;
  givenName: string | null;
  surname: string | null;
  mail: string | null;
  userPrincipalName: string | null;
  department: string | null;
  jobTitle: string | null;
  accountEnabled: true;
  email: string;
}

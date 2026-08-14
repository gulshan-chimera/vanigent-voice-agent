// src/lib/callerLookup.ts
// Given a caller's phone number (as sent by VAPI), checks whether it belongs
// to an employee in Microsoft Entra ID by querying the Graph /users endpoint.

import { getGraphAccessToken } from "./graphAuth";
import { CallerLookupResult, EntraUser } from "../types/vapi";

/**
 * Builds a set of plausible phone number variants to check against Entra ID,
 * since numbers may be stored in different formats by different admins
 * (with/without +, with/without country code, spaced, dashed, etc.)
 */
function buildPhoneVariants(rawNumber: string): string[] {
  const variants = new Set<string>();

  // Original, as received
  variants.add(rawNumber);

  // Strip all non-digit characters except leading +
  const digitsOnly = rawNumber.replace(/[^\d+]/g, "");
  variants.add(digitsOnly);
  console.log('hello')

  // Without leading +
  if (digitsOnly.startsWith("+")) {
    variants.add(digitsOnly.slice(1));
  } else {
    variants.add(`+${digitsOnly}`);
  }

  // Without country code, assuming India (+91) or similar 2-digit code,
  // keeping the last 10 digits — common local storage format
  const last10 = digitsOnly.replace(/\D/g, "").slice(-10);
  if (last10.length === 10) {
    variants.add(last10);
    variants.add(`+91${last10}`);
    variants.add(`91${last10}`);
  }

  return Array.from(variants).filter((v) => v.length > 0);
}

/**
 * Escapes single quotes in a string for safe use inside an OData filter.
 */
function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

export async function lookupCaller(rawNumber: string): Promise<CallerLookupResult> {
  const notAuthenticated: CallerLookupResult = { isAuthenticated: false, user: null };

  if (!rawNumber) {
    console.warn("[CALLER-LOOKUP] No phone number provided — failing closed.");
    return notAuthenticated;
  }

  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[CALLER-LOOKUP] Could not acquire Graph token — failing closed.");
    return notAuthenticated;
  }

  const variants = buildPhoneVariants(rawNumber);

  // Build an OData filter checking mobilePhone and businessPhones against
  // every variant of the number.
  const filterParts: string[] = [];
  for (const variant of variants) {
    const escaped = escapeODataValue(variant);
    filterParts.push(`mobilePhone eq '${escaped}'`);
    filterParts.push(`businessPhones/any(p:p eq '${escaped}')`);
  }
  const filter = filterParts.join(" or ");

  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(
    filter
  )}&$select=id,displayName&$count=true`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual", // required for advanced OData queries like $count and certain filters
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(
        `[CALLER-LOOKUP] Graph API error (${response.status}): ${errorText}`
      );
      return notAuthenticated;
    }

    const data = (await response.json()) as { value: EntraUser[] };

    if (!data.value || data.value.length === 0) {
      console.log(`[CALLER-LOOKUP] No match found for number: ${rawNumber}`);
      return notAuthenticated;
    }

    if (data.value.length > 1) {
      console.warn(
        `[CALLER-LOOKUP] Multiple matches (${data.value.length}) found for number: ${rawNumber}. Using first match.`
      );
    }

    const matchedUser = data.value[0];
    console.log(
      `[CALLER-LOOKUP] Match found for ${rawNumber}: ${matchedUser.displayName} (${matchedUser.id})`
    );

    return { isAuthenticated: true, user: matchedUser };
  } catch (error) {
    console.error(
      `[CALLER-LOOKUP] Network error during Graph query: ${(error as Error).message}`
    );
    return notAuthenticated;
  }
}
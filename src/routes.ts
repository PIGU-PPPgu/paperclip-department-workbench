export function officePath(companyPrefix: string | null | undefined, pathname: string, query = ""): string {
  const prefix = companyPrefix || pathname.split("/").filter(Boolean)[0] || "";
  return `${prefix ? `/${prefix}` : ""}/office${query}`;
}

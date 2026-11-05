// Sorek Register - settings. Change values here only.
window.SOREK_CONFIG = {
  clientId: "26c42f8f-f9b3-4e84-9a38-d3d7d25c5493",   // Entra: Application (client) ID
  tenantId: "561f9c1d-a9d7-418b-a005-a5010ff778ff",   // Entra: Directory (tenant) ID
  siteHost: "sorekenergy.sharepoint.com",
  sitePath: "/sites/SorekEnergy",
  lists: { register: "EmailRegister", companies: "Companies", keywords: "Keywords" },
  project: "SOR",
  issuer: "SRK",
  generalMailbox: "sorek@sorek.energy",
  // Domains of Sorek staff (emails from these are "outgoing")
  ownDomains: ["sorek.energy", "edeltech.co.il"],
  // Values that exist in the EmailRegister "Company" choice column; any other code is written as "Other"
  companyChoices: ["GOV", "FIC", "SNR", "SIE"]
};

/* Configuration Azure AD - a completer apres l'inscription de l'application.
   Voir GUIDE-INSTALLATION.md pour la marche a suivre complete. */
window.DEAL_FLOW_CONFIG = {
  // Colle ici le "Application (client) ID" recupere dans Microsoft Entra > Inscriptions d'applications.
  CLIENT_ID: "7a925922-25d3-4dee-823f-7e6b1eb2f11b",

  // Ne pas modifier : restreint la connexion aux comptes Microsoft personnels.
  AUTHORITY: "https://login.microsoftonline.com/consumers",

  // Nom du fichier stocke dans le dossier applicatif de ton OneDrive (Apps/Deal Flow Tracker/).
  STATE_FILENAME: "deal-flow-tracker-state.json",
};

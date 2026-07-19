/* Configuration Azure AD - a completer apres l'inscription de l'application.
   Voir GUIDE-INSTALLATION.md pour la marche a suivre complete. */
window.DEAL_FLOW_CONFIG = {
  // Colle ici le "Application (client) ID" recupere dans Microsoft Entra > Inscriptions d'applications.
  CLIENT_ID: "6ccea910-1bdd-46d8-99e7-e8abb97e6311",
  // Ne pas modifier : restreint la connexion aux comptes Microsoft personnels.
  AUTHORITY: "https://login.microsoftonline.com/consumers",

  // Nom du fichier stocke dans le dossier applicatif de ton OneDrive (Apps/Deal Flow Tracker/).
  STATE_FILENAME: "deal-flow-tracker-state.json",
};

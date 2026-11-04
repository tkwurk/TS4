// Run this function to wipe ALL properties if you want a clean slate
function forceDeleteAllProperties() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log("All properties completely cleared.");
}
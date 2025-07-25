// module.exports = {
//   auth: {
//     clientId: process.env.CLIENT_ID,
//     authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
//     clientSecret: process.env.CLIENT_SECRET,
//   },
//   system: {
//     loggerOptions: {
//       loggerCallback(loglevel, message) {
//         console.log(message);
//       },
//       piiLoggingEnabled: false,
//       logLevel: 3,
//     },
//   },
// };



module.exports = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, // This will use 'consumers'
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: 3,
    },
  },
};

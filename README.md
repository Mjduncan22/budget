# Budget (web app)

The same Budget app as the Android version, as a website that installs to any phone's home screen
(iPhone: Safari > Share > Add to Home Screen; Android: Chrome > ⋮ > Add to Home screen).

* Sign in with your own Microsoft account; the app reads and updates **your** `Budget Data.xlsx` in OneDrive.
  Michael's: `Documents/Finances/Budget/Budget Data.xlsx`. Lena's: `Lena Budget/Budget Data.xlsx`.
* No data is stored on this site. It's plain HTML/JS hosted on GitHub Pages; your data only travels
  between your phone, Microsoft (sign-in + OneDrive) and CoinGecko (crypto prices).
* To publish a change: edit the files, bump `VERSION` in `sw.js` and `js/app.js`, commit and push.
  Phones pick up the new version the next time the app opens.

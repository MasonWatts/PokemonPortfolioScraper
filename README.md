# PokemonPortfolioScraper
Several data scraping tools that estimate market value of elements of your Pokemon card collection.

Graded cards: Uses pokemonprice.com, a repository of sold prices of graded pokemon cards (sometimes referred to as "slabs" in the colector community), to produce a spreadsheet of current market value for your graded card collection (input via csv file).

Ungraded cards: Pulls from pokellector.com, a collection tracker for un-graded cards, to catalogue your current collection. Then uses tcgplayer.com, a card marketplace, to tabulate current market value of your un-graded collection in a spreadsheet, organized by released card set.

Theme decks: Uses bulbapedia.com, a Pokemon knowledge wiki, to catalogue the cards contained within a given Theme deck (products released by Pokemon with fixed contents). Then uses tcgplayer.com, a card marketplace, to tabulate the current market value of the deck's contents if they were sold inividually. [For a while, some theme decks were selling on eBay for less than their contained cards were worth]

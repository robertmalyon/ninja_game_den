# Ninja Game Den

## Disclaimer

This was a big project that I had to scrap. It was my first proper project, I was in at the deep end. I'm proud of it because of how I came to it with little knowledge and achieved some things I'm pleased with. BUT, it's very rough indeed. I was highly inexperienced and I can see lots of things that should have been done differently based on where I'm at now. But I've left it mostly 'as-is'. VS Code's extensions have made a few improvements, but it is mostly in its raw form. Full story below.

## The Long Read

Deep breath... In 2017 I opened a small shop on London Road in Brighton called Ninja Game Den. The business was a partnership and I owned half of the company. We sold pre-owned games, consoles and accessories specializing in 'retro' games.

The technology underpinning our operation was quite primitive, but we didn't have a lot of spare money to invest. One of the biggest problems we faced was the fact that we bought all our stock from our customers. Our items were all pre-owned and our EPOS system needed to facilitate trade-ins as well as sales and returns and all the usual stuff. The only EPOS system that offered this functionality, but which was an 'off the shelf' solution as opposed to a prohibitively expensive bespoke system, was quite an old system based on an MS Access database. We ran this for a while, we had it connected to a good label printer and we had a little software program written by a developer to connect the printer to the EPOS system and allow automated printing of price labels and barcodes.
After a time we found that we wanted to expand into selling online and we also came to find that a lot of our internal processes were highly labour intensive. Our current software was no longer able to keep up but our finances prevented us from affording an upgrade.
So I decided to try and make software...

My background in coding was not very strong at this point. I went to Brighton University in 2005 and studied Computer Science for two years but dropped out without completing my degree. I studied some Java and we also dabbled in web design, both of which proved helpful, but also seriously out of date and seriously hard to remember given about 15 years had passed! I also remembered that one of the reasons I failed level two was that I dedicated about 90% of my independent study time to work on my 3D Graphics and Animation module using 3DS Max to create a 3D model of a PSP, which I was super happy with. The University was so impressed that they used it on their website as an example of outstanding work. I was able to proudly point this work out to people for about 5 years after I dropped out before it was eventually replaced.

Anyway, back to the slightly more modern-day. Some online research pointed to Python as a good programming language to learn for lots of reasons I won't go too deeply into right now.
I created a company website using WordPress and installed the WooCommerce plugin. My next aim was to somehow connect the in-store EPOS system's stock levels with the website's stock levels to automate the process of uploading products and to make sure that we wouldn't sell something online and then accidentally sell it instore before it was dispatched. Or end up with unsynchronised stock madness leading to frequent stock counts of our 2000 odd inventory.

Experienced developers will know that this is quite a big job, especially when you didn't finish university and haven't coded for 15 years. I genuinely went a little bit mad I think, I was simultaneously learning to code in a haphazard ad-hoc directionless job-by-job way whilst implementing this new system that I hadn't planned thoroughly and continuously evolved. And I was doing it between the hours of about 10 pm and 1 am because the shop was open 7 days a week and we had no staff and we normally worked for between 60 and 84 hours a week just to do all the normal stuff.

I never finished it.

But I did achieve a few things. There were a few moments where I experienced a sense of achievement that was euphoric. I was quite addicted to the buzz of hitting milestones. Sadly, I didn't know anyone who could appreciate or understand what I had done so I couldn't share it without boring people. But if just one person sees this and thinks that, in the context of the circumstances I've outlined above, there are some clever little ideas there, then that would make me happy because I sunk a lot of hours into this and it remains unfinished and unloved by anyone except me.

For me, it represents a turning point and a 'spring-board' that kick-started my current goal which is to become a 'full-stack developer' by following a much more formal, goal-oriented learning path drenched in the humility of having tried to do something too big for me, but richer in knowledge for having tried!

## Things I did Manage to Achieve

I made a connection to the MS Access database that powered the EPOS system. This allowed me to retrieve, edit, add replace etc product data.

I created a decent looking GUI using Tkinter which would allow a barcode to be scanned and present product details from the database.

The GUI also allowed changes to be made to the product, which could then be written back to the database.

The GUI would allow the creation of a product price label and barcode, which could them be printed on our Zebra label printer.

When printing a product's labels the option to send the product details to the company website is offered. This functionality works and it sends a POST request to the WooCommerce API creating a new product in the online database.

Finally the program 'hijacked' an unused table in the EPOS system's database and allowed the tracking of unique products as opposed to generic ones. This assigned a BAT number to a product and printed the BAT number as a barcode to cover a games built-in manufacturer barcode. This enabled us to include details about the individual product's condition.

One additional feature that I was pleased with was that I used Python's Regular Expressions module to improve the presentation and format of data taken from the database. So, often products had been entered as block capitals with internal notes added in brackets. this didn't matter as only part of the details appeared on a receipt and it was just for our records. When it came to using these titles to populate the website I ensured that titles were set to 'title case' and additional details included in brackets were removed, whitespace stripped and other small changes. I also included some default values for some of the fields to speed up data input.
Further Ideas I had

Apart from deploying the thing I did have a few other goals that I also didn't get to!
I was working on a way of using Python to automate the naming of IMG files so that we could take product photos and have them tie in with the relevant product based on order so that they could be bulk-uploaded.

I was working on a way of checking and updating our prices using web scraping.

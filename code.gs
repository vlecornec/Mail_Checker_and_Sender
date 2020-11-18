function getGeocodingRegion() {
  return PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'fr';
}


function Deg2Rad(deg) {
  return deg * Math.PI / 180;
}

function PythagorasEquirectangular(lat1, lon1, lat2, lon2) {
  lat1 = Deg2Rad(lat1);
  lat2 = Deg2Rad(lat2);
  lon1 = Deg2Rad(lon1);
  lon2 = Deg2Rad(lon2);
  var R = 6371; // km
  var x = (lon2 - lon1) * Math.cos((lat1 + lat2) / 2);
  var y = (lat2 - lat1);
  var d = Math.sqrt(x * x + y * y) * R;
  return d;
}

var cities = [
["Aix-en-Provence", 43.5281228, 5.4478879],
["Amiens", 49.8932679, 2.2954923],
["Angers", 47.4702484, -0.5479967],
["Annecy", 45.9033057, 6.12980935],
["Antibes-Cannes", 43.5531397, 7.0174563],
["Avignon", 43.9436123, 4.8062061],
["Belfort - Montbeliard", 47.6384355, 6.86237995],
["B&eacute;ziers - Narbonne", 43.3428151, 3.21342515],
["Bordeaux", 44.8413806, -0.577115],
["Brest - Quimper", 48.3914598, -4.4848262],
["Caen", 49.1810607, -0.3627764],
["Chamb&eacute;ry", 45.6925223, 5.9094329],
["Clermont-Ferrand", 45.7798396, 3.0865628],
["Colmar", 48.0794311, 7.3580665],
["Dijon - Beaune", 47.3210156, 5.0393361],
["Essonne", 48.6559568, 2.4085267],
["Fr&eacute;jus - Saint-Rapha&euml;l", 43.4235909, 6.7696424],
["Grenoble", 45.1851489, 5.7262285],
["Hauts-de-Seine", 48.9030336, 2.304768],
["La Baule-Saint Nazaire", 47.2737696, -2.2142795],
["La Rochelle", 46.1586365, -1.1515646],
["Le Havre", 49.49429, 0.1141],
["Le Mans", 47.9952925, 0.1968826],
["Lille", 50.6313919, 3.0596752],
["Limoges", 45.832172, 1.2603562],
["Lorient - Vannes", 47.6606103, -2.7624331],
["Lyon", 45.7629354, 4.843525],
["Marseille", 43.29225295, 5.38690575],
["Metz", 49.1123514, 6.1776342],
["Montpellier", 43.6052092, 3.8806329],
["Mulhouse", 47.7477981, 7.3361225],
["Nancy", 48.69153915, 6.1825616],
["Nantes", 47.21739835, -1.5537669],
["Nice", 43.7001698, 7.2657269],
["N&icirc;mes - Al&egrave;s", 43.836699, 4.360054],
["N&icirc;mes - Al&egrave;s", 44.133333, 4.083333],
["Oise", 49.1995857, 2.5777531],
["Orl&eacute;ans", 47.902757, 1.908842],
["Paris", 48.8631399, 2.3447469],
["Pays Basque", 43.4925781, -1.4751586],
["Pays Bearnais", 43.3044119, -0.3715582],
["Perpignan", 42.6966774, 2.8909497],
["Poitiers", 46.5801022, 0.3412419],
["Reims", 49.2580803, 4.0317017],
["Rennes", 48.1055679, -1.6759291],
["Rouen", 49.44174765, 1.09104625],
["Saint-Etienne", 45.4396383, 4.38910285],
["Seine et Marne", 48.9587474, 2.8833628],
["Strasbourg", 48.5839424, 7.74649105],
["Toulon", 43.1202542, 5.93465615],
["Toulouse", 43.6069831, 1.441872],
["Tours", 47.39088005, 0.69083505],
["Val-de-Marne", 48.7984191, 2.3999118],
["Val-d-Oise", 48.94621215, 2.24659215],
["Valence - Montelimar", 44.933393, 4.89236],
["Valence - Montelimar", 44.556944, 4.749496],
["Valenciennes", 50.3574945, 3.5228022],
["Yvelines", 48.8040787	, 2.12620385],
];

function NearestCity(latitude, longitude) {
  var mindif = 99999;
  var closest;

  for (index = 0; index < cities.length; ++index) {
    var dif = PythagorasEquirectangular(latitude, longitude, cities[index][1], cities[index][2]);
    if (dif < mindif) {
      closest = index;
      mindif = dif;
    }
  }
 var division= cities[closest][0];
 return division
}

function countRepitition(searchword) {

   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var allappross = ss.getSheetByName('All Appros');
   var allaprosEndRow = allappross.getLastRow();

   var count = 0;
   var values = allappross.getRange(2, 3, allaprosEndRow, 3).getValues(); //Start of Range (5,1) -> A5, End of Range (9,2) -> B9

   for (var row in values) {
     for (var col in values[row]) {

       if (values[row][col] == searchword) {
         count++;
       }

     }
   }

  return count
}
function checkdivisionviamail()
{
  // Use sheet
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var checkss= ss.getSheetByName('Check')
  var current_date = new Date();
  var weekday = current_date.getDay();
  if ((weekday != 0) && (weekday != 6)){
  var allappross = ss.getSheetByName('All Appros');
  var rejectss = ss.getSheetByName('Reject');
  var rejectStartRow = rejectss.getRange("E2").getRow();
  var rejectnEndRow = rejectss.getLastRow();
  // Gmail query
  var label = GmailApp.getUserLabelByName("Approval Request/A Traiter");
  var label2= GmailApp.getUserLabelByName("Approval Request/Check Division Fait");
  var label3= GmailApp.getUserLabelByName("Approval Request/Rejet pour Division");
  var label4= GmailApp.getUserLabelByName("Approval Request");
  var label5= GmailApp.getUserLabelByName("Approval Request/Multiadresse");
  var label6= GmailApp.getUserLabelByName("Approval Request/Appro Déléguée");
  var threads = label.getThreads();
    for (var i=0; i<threads.length; i++)
    {
      var messages = threads[i].getMessages();

      for (var j=0; j<messages.length; j++)
      {
        // Find data
        var from = messages[j].getFrom();
        var time = messages[j].getDate();
        var sub = messages[j].getSubject();
        var reply=messages[j].getReplyTo();
        if (reply==""){
                        threads[i].removeLabel(label);
                        threads[i].addLabel(label6);
                        threads[i].addLabel(label4);
          // pas de mail salesforce à qui répondre
        }
        else{
          var idmail=messages[j].getId();
          var treated=countRepitition(idmail);
          if (treated>0){
           }
          else{
          var msge=messages[j].getBody();
          var subjects=messages[j].getSubject()
          if (subjects.indexOf("déléguée")>-1||subjects.indexOf("Delegate")>-1){
           threads[i].removeLabel(label);
           threads[i].addLabel(label6);
           GmailApp.markThreadRead(threads[i]);
          }
          else{
          var begdiv= msge.indexOf("Division:");
          var enddiv=msge.indexOf("<br />", begdiv);
          var division = msge.substring(begdiv+10,enddiv).trim();
          var endoppty=msge.indexOf("You can view the Opportunity here</a>");
          var begoppty=endoppty-36;
          var oppty= msge.substring(begoppty,begoppty+18);
          var begadresse=msge.indexOf('<b>Redemption Address(es):</b>');
          var endadresse=msge.indexOf('</table>',begadresse);
          var adresse = msge.substring(begadresse,endadresse);
          Logger.log(adresse);
          var numberofadress=(adresse.split("tr>").length-3)/2;
          Logger.log(numberofadress);
          if (numberofadress==1)
           {
              var begstreetline1=adresse.indexOf('<td style=');
              var begstreetline2=adresse.indexOf('<td style=',begstreetline1+1);
              var begzipcode=adresse.indexOf('<td style=',begstreetline2+1);
              var endzipcode=adresse.indexOf('</td>',begzipcode+1);
              var begstate=adresse.indexOf('<td style=',begzipcode+1);
              var begcountry=adresse.indexOf('<td style=',begstate+1);
              var endcountry=adresse.indexOf('</td>',begcountry+1);
              var country=adresse.substring(begcountry+53,endcountry).trim();
              Logger.log(country);
              if (country!='FR'){
              }
              else{
              var streetline1=adresse.substring(begstreetline1+53,begstreetline2-15).trim();
                var streetline2=adresse.substring(begstreetline2+53,begzipcode-15).trim();
                var zipcode=adresse.substring(begzipcode+53,endzipcode).trim();
                Logger.log(zipcode);
                var department=zipcode.substr(0,zipcode.length-3);
                if (department==75)
                {
                  var divisiontoput="Paris";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                 }
                 else if(department==60)
                {
                  var divisiontoput="Oise";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                 }
                 else if(department==77)
                {
                  var divisiontoput="Seine et Marne";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                 }
                else if(department==78)
                {
                  var divisiontoput="Yvelines";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                 }
                else if(department==91)
                {
                  var divisiontoput="Essonne";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                }
                else if(department==92)
                {
                  var divisiontoput="Hauts-de-Seine";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                  }
                else if(department==93)
                {
                  var divisiontoput="Paris";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                  }
                else if(department==94)
                {
                  var divisiontoput="Val-de-Marne";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                  }
                else if(department==95)
                {
                  var divisiontoput="Val-d-Oise";
                  var address="undefined";
                  var lat="undefined";
                  var lon="undefined";
                }
                else {
                  var state=adresse.substring(begstate+53,begcountry-15).trim();
                  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
                  var location;
                  var address =streetline1+", "+zipcode+", "+state+", "+country;
                  address = address.replace(/%27/g,"'").replace(/&ocirc;/g,"ô").replace(/&egrave;/g,"è").replace(/&eacute;/g,"é").replace(/&euml;/g,"ë").replace(/&Ocirc;/g,"Ô").replace(/&Egrave;/g,"È").replace(/&Eacute;/g,"É").replace(/&euml;/g,"Ë")
                  var address2=address.toString().replace(/'/g,"%2F").replace(/ô/g,"%C3%B5").replace(/é/g,"%C3%A9").replace(/è/g,"%C3%A8").replace(/à/g,"%C3%A0").replace(/ê/g,"%C3%AA").replace(/É/g,"%C3%89").replace(/Î/g,"%C3%8E").replace(/î/g,"%C3%AE");
                  var addresse3="Zipcode "+zipcode+", "+country;
                  location= geocoder.geocode(addresse3);
                    if (location.status == 'OK')
                    {
                      var lat = location["results"][0]["geometry"]["location"]["lat"];
                      var lon = location["results"][0]["geometry"]["location"]["lng"];
                      var divisiontoput=NearestCity(lat, lon)
                    }
                    else
                    {
                      var lat="undefined";
                      var lon="undefined";
                      var divisiontoput="undefined";
                    }

                 }
                 allappross.appendRow([from, time,idmail, sub, reply,division,oppty,numberofadress,address,lat,lon,divisiontoput,country,zipcode])
                 if (divisiontoput!==division && divisiontoput!=="undefined")
                      {
                        var msgtosend="<p>No</p>" + "<p> Hello, merci de changer la division en " + divisiontoput + " stp.</p>"
                        MailApp.sendEmail({
                            to: reply,
                            subject: "Opportunity reject"+" "+oppty,
                            htmlBody: "No<br/>" +
                            "Hello, merci de changer la division en " + divisiontoput + " stp."
                          })
                        rejectss.appendRow([from, time,idmail, sub, reply,division,oppty,numberofadress,address,lat,lon,divisiontoput,msgtosend,address2,location,zipcode]);
                        GmailApp.markThreadRead(threads[i])
                        threads[i].removeLabel(label);
                        threads[i].addLabel(label3);
                        threads[i].addLabel(label4);

                         }
                     else
                     {
                        checkss.appendRow([from, time,idmail, sub, reply,division,oppty,numberofadress,address,lat,lon,divisiontoput,country,zipcode]);
                        threads[i].removeLabel(label);
                        threads[i].addLabel(label2);
                        threads[i].addLabel(label4);
                       }
                 }
                 }
             else {
                        threads[i].removeLabel(label);
                        threads[i].addLabel(label4);
                        threads[i].addLabel(label5);// Plusieurs Redemption
               }
             }
          }
          }
          }
        }
    }
}
  
   function test(){
   MailApp.sendEmail({
                          to: "Email Approval <0-4hbb0t44qbts1v.nyce6v9ac2j7594b@7n9oq2xq1z5kfc6k.lzfdoyt.8-khmyeac.na8.wrkfl.salesforce.com>",
                          subject: "Opportunity reject"+"test",
                          htmlBody: "No<br />" + "Hello, merci de changer la division en Fr&eacute;jus - Saint-Rapha&euml;l stp." 
                        }) 
   }
   
   function test2(){
   var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
                var location;
                var addresse="30250"
                var addresse2=addresse.replace(/'/g,"%2F").replace(/ô/g,"%C3%B5").replace(/é/g,"%C3%A9").replace(/è/g,"%C3%A8").replace(/à/g,"%C3%A0").replace(/ê/g,"%C3%AA").replace(/É/g,"%C3%89").replace(/Î/g,"%C3%8E").replace(/î/g,"%C3%AE");
                var addresse3="Zipcode "+addresse+", "+"FR";
               location= geocoder.geocode(addresse2);
                    var lat = location["results"][0]["geometry"]["location"]["lat"];
                    var lon = location["results"][0]["geometry"]["location"]["lng"];
                    var divisiontoput=NearestCity(lat, lon)  
    Logger.log(lat + " "+lon + " "+divisiontoput)
    }
               

// Emails of current execs
const emailPresident = 'alexis.demetriou@mail.mcgill.ca';
const emailVPinternal = ''


/**
 * @author: Andrey S Gonzalez
 * @date: Nov 13, 2023
 * @update: Sep 24, 2024
 * 
 * Returns the headrunners' emails for input `headrun`.
 */ 
function getHeadRunnerEmail(headrun) {

  // Head Runner Emails
  const abigailFinch = 'abigail.finch@mail.mcgill.ca';
  const aidenLee = 'jihong.lee@mail.mcgill.ca';
  const alexanderHebiton = 'alexander.hebiton@mail.mcgill.ca';
  const ameliaRilling = 'amelia.rilling@mail.mcgill.ca';
  const bridgetAndersen = 'bridget.andersen@mail.mcgill.ca';
  const camilaCognac = 'camila.cognac@mail.mcgill.ca';
  const charlesVillegas = 'charles.villegas@mail.mcgill.ca';
  const edmundPaquin = 'edmund.paquin@mail.mcgill.ca';
  const emersonDarling = 'emerson.darling@mail.mcgill.ca';
  const filipSnitil = 'filip.snitil@mail.mcgill.ca';
  const bellaVignuzzi = 'isabella.vignuzzi@mail.mcgill.ca';
  const jamesDiPaola = 'james.dipaola@mail.mcgill.ca';
  const julietteAdeline = 'juliette.adeline@mail.mcgill.ca';
  const justinCote = 'justin.cote2@mail.mcgill.ca';
  const kateRichards = 'katherine.richards@mail.mcgill.ca';
  const lakshyaSethi = 'lakshya.sethi@mail.mcgill.ca';
  const liamGrant = 'liam.grant@mail.mcgill.ca';
  const liamMurphy = 'liam.murphy3@mail.mcgill.ca';
  const madisonHughes = 'madison.hughes@mail.mcgill.ca';
  const michaelRafferty = 'michael.rafferty@mail.mcgill.ca';
  const nicolasMorrison = 'nicolas.morrison@mail.mcgill.ca';
  const pooyaPilehChiha = 'pooya.pilehchiha@mail.mcgill.ca';
  const prabhjeetSingh = 'prabhjeet.singh@mail.mcgill.ca';
  const rachelMattingly = 'rachel.mattingly@mail.mcgill.ca';
  const roriSa = 'rori.sa@mail.mcgill.ca';
  const sophiaLongo = 'sophia.longo@mail.mcgill.ca';
  const tessLedieu = 'tess.ledieu@mail.mcgill.ca';
  const theoGhanem = 'theo.ghanem@mail.mcgill.ca';

  // Head Runners associated to each head run
  const tuesdayHeadRunner = [tessLedieu, julietteAdeline, jamesDiPaola, michaelRafferty, liamMurphy, bridgetAndersen];
  const wednesdayHeadRunner = [kateRichards, nicolasMorrison, sophiaLongo, camilaCognac, alexanderHebiton];
  const thursdayHeadRunner = [charlesVillegas, ameliaRilling, emersonDarling, justinCote, liamGrant];
  const saturdayHeadRunner = [abigailFinch, rachelMattingly, filipSnitil, theoGhanem, bellaVignuzzi, lakshyaSethi];
  const sundayHeadRunner = [prabhjeetSingh, edmundPaquin, roriSa, madisonHughes, pooyaPilehChiha, aidenLee];

  // Easier to decode from input `headrun`
  switch (headrun) {
  case 'TuesdayPM'   : return tuesdayHeadRunner;
  case 'WednesdayPM': return wednesdayHeadRunner;
  case 'ThursdayAM' : return thursdayHeadRunner;
  case 'SaturdayAM': return saturdayHeadRunner;
  case 'SundayPM': return sundayHeadRunner;

  default : return '';
  }

}


/**
 * @author: Andrey S Gonzalez
 * @date: Nov 13, 2023
 * @update: Sep 24, 2024
 * 
 * Returns the head run details as a string using input `headRunDay`
 *  Tuesday   :  6:00pm            
    Wednesday :  6:00pm          
    Thursday  :  7:30am           
    Saturday  :  10:00am
    Sunday    :  6:00pm
 */ 

function getHeadRunString(headRunDay) {
  switch(headRunDay) {
    case 'TuesdayPM'  : return 'Tuesday 6pm';
    case 'WednesdayPM': return 'Wednesday 6pm';
    case 'ThursdayAM' : return 'Thursday 7:30am';
    case 'SaturdayAM' : return 'Saturday 10am';
    case 'SundayPM'   : return 'Sunday 6pm';

  default : return '';
  }

}


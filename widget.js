<script>
/**
 * introPage - the starting page.
 * parseBtn - button on the introPage with event listener to parse data.
 * wrapper - the wrapper in which we are working.
 */
const introPage = document.getElementById( 'intro-page' );
const wrapper = document.getElementById( 'wrapper' );

const parseBtn = document.getElementById( 'parse-btn' );
parseBtn.addEventListener( 'click', initializeParseForm );
let form;


/**
 * The function initializes form for parsing/inserting data
 */
function initializeParseForm() {
    // create the form with the hint
    form = document.createElement( 'form' );
    const hint = document.createElement( 'p' );
    hint.innerHTML = 'Choose the tables with B095 form, B085 form and the destination table:';
    form.append( hint );
  
    // hide the introPage and append the form
    introPage.hidden = true;
    wrapper.append( form );

    // get the names of the sheets of the current spreadsheet
    new Promise((resolve, reject) => {
        google.script.run
            .withSuccessHandler((json) => {
                const data = JSON.parse(json);
                resolve(data);
             } )
            .withFailureHandler( (err) => {
                reject(err);
            } )
            .getSheetNames();
    })
    // when the names of the spreadsheet are received - generate select HTML elements
    // to select the sheet for the B095, B085 forms and destination sheet.
  .then(
    spreadsheetNames => {
        var btn = generateButtonElement( 'Parse' );
        btn.addEventListener( 'click', runScripts );
        
        // add select elements
        form.append( 
            generateSelectElement( spreadsheetNames, 'B095' ),
            generateSelectElement( spreadsheetNames, 'B085' ),
            generateSelectElement( spreadsheetNames, 'dest' ),
            generateSelectElement( spreadsheetNames, 'monthly' ),
            btn
        );

        
        var selectElems = document.querySelectorAll( 'select' );
        M.FormSelect.init( selectElems );
    },
    
    error => {
      console.log( error ); // TODO
    }
  );
}


/**
 * The function runs the getAllData function
 * @param {Event} e The event object to prevent default behavior. 
 */
function runScripts(e) {
    // change value of button
    var btn = e.target;
    btn.innerHTML = 'Processing';
    e.preventDefault();
    
    // get names of the sheets
    var B095 = form.B095.value;
    var B085 = form.B085.value;
    var dest = form.dest.value;
    
    // parse data
    new Promise((resolve, reject) => {  
        google.script.run
            .withSuccessHandler(
                json => {
                    const data = JSON.parse( json );
                    resolve( data );
                }
            )
            .withFailureHandler(
                error => reject( error )
            )
            .getAllData( B085, B095, dest );        
    })
    .then(
        data => {
            wrapper.append( generateResultsSection( data ) );
            btn.remove();
        },
        err => {
            console.log( err );
        }
    )
    // if monthly is set  - add check
    .then(
      () => {
        if ( form.monthly.value !== 'Choose monthly' ) {
            return new Promise((resolve, reject) => {
                google.script.run.handleDuplicatesAndDifferences( form.monthly.value, form.dest.value );
            });
        }
      }
    );

}

/**
 * The function creates the SELECT html element.
 * @param {Array} data The array of options.
 * @returns {HTMLElement} Html element select.
 */
function generateSelectElement(data, label) {
  const select = document.createElement('select');
  select.setAttribute( 'name', label );
  
  // add options
  data.forEach( item => {
    const option = generateOptionElement( item );
    select.append( option );
  } );
  
  // add disabled and selected option as a placeholder
  const firstOption =  generateEmptyOptionElement( 'Choose ' + label );
  select.prepend( firstOption );
  
  return select;
}


/**
 * Generates the option HTML element.
 * @param {String} name Name and innerHTML of the option.
 * @returns {HTMLElement} Option HTML Element.
 */
function generateOptionElement(name) {
    const option = document.createElement( 'option' );
    option.innerHTML = name;
    option.value = name;

    return option;
}


/**
 * Generates the empty option HTML element (selected and disabled).
 * @param {String} name Placeholder for the option element.
 * @returns {HTMLElement} Option HTML Element.
 */
function generateEmptyOptionElement(text) {
    const option = generateOptionElement(text);
    option.disabled = true;
    option.selected = true;

    return option;
}


/**
 * Generates the button with the given text and basic styles.
 * @param {String} text innerHTML for the button.
 * @returns {HTMLElement} Button.
 */
function generateButtonElement(text) {
    const btn = document.createElement('button');

    btn.innerHTML = text;
    btn.classList.add('btn', 'waves-effect', 'waves-light');

    return btn;
}


/**
 * Generates the section with errors and warnings.
 * @param {Array} data Array of objects with records and descriptions.
 * @returns {HTMLElement} Section element with warnings and errors.
 */
function generateResultsSection(data) {
    const section = document.createElement( 'section' );
    
    data.forEach( item => {
        // if item if empty - dismiss
        if ( item.data.length ) {
            const div = document.createElement( 'div' );
            const desc = document.createElement( 'p' );

            // give description
            desc.style.color = '#cc3300';
            desc.innerHTML = item.description;

            // generate ol list
            div.append( desc, generateRecordElements( item.data ) );
            section.append( div );
        }
    } );

    return section;
}


/**
 * The function creates HTML list Element (ol)
 * @param {Array} array Array of records
 * @returns {HTMLElement} List element.
 */
function generateRecordElements(array) {
    const list = document.createElement( 'ol' );

    array.forEach( item => {
        const listItem = document.createElement( 'li' );
        listItem.innerHTML = item.join( ', ' );
        list.append( listItem );
    } );

    return list;
}
</script>

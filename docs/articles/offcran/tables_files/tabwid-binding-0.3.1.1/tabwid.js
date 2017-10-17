HTMLWidgets.widget({

  name: 'tabwid',

  type: 'output',

  factory: function(el, width, height) {

    // TODO: define shared variables for this instance

    return {

      renderValue: function(x) {

        // TODO: code to render the widget, e.g.
        el.innerHTML = x.html;

        var sheet = document.createElement('style');
        sheet.innerHTML = x.css;
        document.body.appendChild(sheet);
        el.style.width = null;
        el.style.height = null;

      },

      resize: function(width, height) {

      }

    };
  }
});

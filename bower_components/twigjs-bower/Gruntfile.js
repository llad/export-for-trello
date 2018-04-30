module.exports = function(grunt) {

  var pkg = grunt.file.readJSON('package.json');

  // Project configuration.
  grunt.initConfig({
    pkg: pkg,
    exec: {
      update: "cd ./twig.js/ && git pull origin master && cd ../",
      build: "cd ./twig.js/ && npm run build && cd ../"
    },
    copy: {
      main: {
        files: [
          {
            cwd: 'twig.js/',
            src: 'twig.*',
            dest: 'twig/',
            expand: true
          }
        ]
      }
    },
    gitadd: {
      dist: {
        options: {
          force: false
        },
        files: {
          src: ['*']
        }
      }
    },
    gitcommit: {
      dist: {
        options: {
          cwd: "./",
          message: "Update to <%= pkg.version %>"
        },
        files: [
          {
            src: ["*", "!node_modules", "!twig.js"],
            expand: true,
            cwd: "./"
          }
        ]
      }
    },
    gittag: {
      dist: {
        options: {
          tag: '<%= pkg.version %>'
        }
      }
    },
    gitpush: {
      dist: {
        options: {
          remote: 'origin'
          // Target-specific options go here. 
        }
      }
    },
  });

  // Load the plugin that provides the "uglify" task.
  require('load-grunt-tasks')(grunt);

  grunt.registerTask('test', function() {
    var twig = grunt.file.readJSON('twig.js/package.json');

    

    if(pkg.version !== twig.version){
      pkg.version = twig.version;
      var fs = require("fs"); 
      var json = JSON.stringify(pkg, null, 2); 
      fs.writeFileSync("package.json",json);

      grunt.config('pkg', twig.version);
      
    }

  });

  // Default task(s).
  grunt.registerTask('default', ['exec:update', 'exec:build', 'copy', 'test']);

};
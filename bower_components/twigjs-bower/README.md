# twigjs-bower
A bower package for Twig js


Since the main twigjs project no longer supports bower, this has been created to plug that gap.

The files are build from the main project.

To use:
```
bower install twigjs-bower --save
```

## A note on browser usage; the current docs say:

```javascript
var template = twig({
    data: 'The {{ baked_good }} is a lie.'
});

console.log(
    template.render({baked_good: 'cupcake'})
);
// outputs: "The cupcake is a lie."
```

However this is resulting in `twig` being undefined.

Instead, this will work:

```javascript
var template = Twig.twig({
    data: 'The {{ baked_good }} is a lie.'
});

console.log(
    template.render({baked_good: 'cupcake'})
);
// outputs: "The cupcake is a lie."
```

All other documentation is as per the [main project](https://github.com/twigjs/twig.js/wiki).
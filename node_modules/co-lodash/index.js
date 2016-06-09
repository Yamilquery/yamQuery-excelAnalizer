var _ = module.exports = require('lodash');

function makeCallback(callback, thisArg) {
    return _.isGenerator(callback) ?
        _.isUndefined(thisArg) ? callback : callback.bind(thisArg) :
        function *() { yield callback.apply(thisArg, arguments); };
}

var eachImpl = function *(collection, callback) {
    var index = -1,
        length = collection ? collection.length : 0;

    if (typeof length == 'number') {
        while (++index < length && false !== (yield* callback(collection[index], index, collection)))
            ;
    } else {
        var ownProps = _.keys(collection),
            length = ownProps ? ownProps.length : 0;

        while (++index < length && false !== (yield* callback(collection[ownProps[index]], ownProps[index], collection)))
            ;
    }
    return collection;
};

_.mixin({
    isGenerator: function (obj) {
        return obj && obj.constructor && obj.constructor.name === 'GeneratorFunction';
    },

    sleep: function (mill) {
        return function (cb) { setTimeout(cb, mill); };
    },

    makeAsync: function (cb, thisArg) {
        return function () {
            var args = arguments;
            process.nextTick(function () { cb.apply(thisArg, args); });
        };
    },

    coEach: function* (collection, callback, thisArg) {
        return yield* eachImpl(collection, makeCallback(callback, thisArg));
    },

    coMap: function* (collection, callback, thisArg) {
        callback = makeCallback(callback, thisArg);

        var ret = [];
        yield* eachImpl(collection, function *(item, index, coll) { ret.push(yield* callback(item, index, coll)); });
        return ret;
    },

    coReduce: function* (collection, callback, accumulator, thisArg) {
        if (!collection) return accumulator;
        var noaccum = arguments.length < 3;
        callback = makeCallback(callback, thisArg);

        if (noaccum && _.isArray(collection)) accumulator = collection[0];

        yield* eachImpl(collection, function *(item, index, coll) {
            accumulator = yield* callback(accumulator, item, index, coll);
        });

        return accumulator;
    }
});

var co = require('co'),
    _ = require('../');

co(function *() {
    yield* _.coEach([1,2,3,4,5], function* (v) {
        console.log('each:', v);
        yield _.sleep(1000);
    });

    console.log(yield* _.coMap([1,2,3,4,5], function *(v) {
        console.log('map:', v);
        yield _.sleep(1000);
        return v * 2;
    }));

    console.log(yield* _.coReduce([1,2,3,4,5], function*(m, v) {
        console.log('reduce:', v);
        yield _.sleep(1000);
        return m + v;
    }, 0));
});

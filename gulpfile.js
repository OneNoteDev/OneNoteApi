/// <binding BeforeBuild='build' />
"use strict";
var require;

var browserify = require("browserify");
var del = require("del");
var fs = require("file-system");
var globby = require("globby");
var gulp = require("gulp");
var header = require("gulp-header");
var merge = require("merge-stream");
var plumber = require("gulp-plumber");
var qunit = require("node-qunit-phantomjs");
var rename = require("gulp-rename");
var runSequence = require("run-sequence");
var source = require("vinyl-source-stream");
var ts = require("gulp-typescript");
var tslint = require("gulp-tslint");
var tsd = require("gulp-tsd");
var uglify = require("gulp-uglify");

var PATHS = {
	SRCROOT: "src/",
	BUILDROOT: "build/",
	BUNDLEROOT: "build/bundles/",
	TARGETROOT: "target/",
	DEFINITIONS: "src/definitions/",
	WEBROOT: "serverRoot/"
};

////////////////////////////////////////
// SETUP
////////////////////////////////////////
gulp.task("cleanDefinitions", function (callback) {
	return del([
		PATHS.DEFINITIONS
	], callback);
});

gulp.task("definitions", function (callback) {
	tsd({
		command: "reinstall",
		config: "tsd.json"
	}, callback);
});

gulp.task("setup", function (callback) {
	runSequence(
		"cleanDefinitions",
		"definitions",
		callback);
});

////////////////////////////////////////
// CLEAN
////////////////////////////////////////
gulp.task("clean", function (callback) {
	return del([
		PATHS.BUILDROOT,
		PATHS.BUNDLEROOT,
		PATHS.TARGETROOT,
		PATHS.WEBROOT
	], callback);
});

////////////////////////////////////////
// COMPILE
////////////////////////////////////////
gulp.task("compile", function() {
	var tsResult = gulp.src([PATHS.SRCROOT + "**/*.+(ts|tsx)", "!**/*.d.ts"])
		.pipe(ts({
			declaration: true,
			inlineSourceMap: true,
			inlineSources: true,
			module: "commonjs"
		}));

	return merge([
		tsResult.dts.pipe(gulp.dest(PATHS.BUILDROOT)),
		tsResult.js.pipe(gulp.dest(PATHS.BUILDROOT))
	]);
});

////////////////////////////////////////
// TSLINT
////////////////////////////////////////
gulp.task("tslint", function () {
	var tsErrorReport = tslint.report("prose", {
		emitError: false,
		reportLimit: 50
	});

	var tsFiles = [PATHS.SRCROOT + "**/*.ts", PATHS.SRCROOT + "**/*.tsx", "!" + PATHS.SRCROOT + "**/*.d.ts"];

	return gulp.src(tsFiles)
		.pipe(plumber())
		.pipe(tslint())
		.pipe(tsErrorReport);
});

////////////////////////////////////////
// BUNDLE
////////////////////////////////////////
function getLicense() {
	return [
		"/*",
		fs.readFileSync("LICENSE", "utf8"),
		"*/"
	].join("\n");
}

gulp.task("bundleApi", function () {
	return browserify(PATHS.BUILDROOT + "scripts/oneNoteApi.js", { standalone: "OneNoteApi" })
		.bundle()
		.pipe(source("oneNoteApi.js"))
		.pipe(header(getLicense()))
		.pipe(gulp.dest(PATHS.BUNDLEROOT));
});

gulp.task("bundleDefinition", function () {
	gulp.src(PATHS.BUILDROOT + "scripts/onenoteApi.d.ts")
		.pipe(gulp.dest(PATHS.BUNDLEROOT));
});

gulp.task("bundleTests", function () {
	return globby.sync(["**/*.js"], { cwd: PATHS.BUILDROOT + "tests" }).map(function (filePath) {
		return browserify(PATHS.BUILDROOT + "tests/" + filePath, { debug: true })
			.bundle()
			.pipe(source(filePath))
			.pipe(gulp.dest(PATHS.BUNDLEROOT + "tests"));
	});
});

gulp.task("bundle", function (callback) {
	runSequence(
		"bundleApi",
		"bundleTests",
		callback);
});

////////////////////////////////////////
// MINIFY BUNDLED
////////////////////////////////////////
gulp.task("minifyBundled", function (callback) {
	var targetDir = PATHS.BUNDLEROOT;

	var minifyTask = gulp.src(PATHS.BUNDLEROOT + "oneNoteApi.js")
		.pipe(uglify({
			preserveComments: "license"
		}))
		.pipe(rename({ suffix: ".min" }))
		.pipe(gulp.dest(targetDir));

	return merge(minifyTask);
});

////////////////////////////////////////
// EXPORT
////////////////////////////////////////
gulp.task("exportApi", function () {
	var modulesTask = gulp.src(PATHS.BUILDROOT + "scripts/**/*.js", { base: PATHS.BUILDROOT + "/scripts" })
		.pipe(gulp.dest(PATHS.TARGETROOT + "modules/"));

	var copyTask = gulp.src([
		PATHS.BUNDLEROOT + "oneNoteApi.js",
		PATHS.BUNDLEROOT + "oneNoteApi.min.js",
		PATHS.SRCROOT + "oneNoteApi.d.ts"
	]).pipe(gulp.dest(PATHS.TARGETROOT));

	return merge(modulesTask, copyTask);
});

gulp.task("exportTests", function () {
	var targetDir = PATHS.TARGETROOT + "tests/";

	var libsTask = gulp.src([
		"node_modules/qunitjs/qunit/qunit.+(css|js)",
		PATHS.SRCROOT + "tests/bind_polyfill.js"
	]).pipe(gulp.dest(targetDir + "libs"));

	var testsTask = gulp.src(PATHS.BUNDLEROOT + "tests/**", { base: PATHS.BUNDLEROOT + "tests/" })
		.pipe(gulp.dest(targetDir));

	var indexTask = gulp.src(PATHS.SRCROOT + "tests/index.html")
		.pipe(gulp.dest(targetDir));

	return merge(libsTask, testsTask, indexTask);
});

gulp.task("export", function (callback) {
	runSequence(
		"exportApi",
		"exportTests",
		callback);
});

////////////////////////////////////////
// RUN
////////////////////////////////////////
gulp.task("runTests", function () {
	return qunit(PATHS.TARGETROOT + "tests/index.html");
});

////////////////////////////////////////
// SHORTCUT TASKS
////////////////////////////////////////
gulp.task("build", function (callback) {
	runSequence(
		"compile",
		"bundle",
		"minifyBundled",
		"export",
		"tslint",
		"runTests",
		callback);
});

gulp.task("full", function (callback) {
	runSequence(
		"clean",
		"build",
		callback);
});

gulp.task("default", ["build"]);

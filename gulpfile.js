/// <binding BeforeBuild='build' />
"use strict";
var require;

var browserify = require("browserify");
var del = require("del");
var forever = require("forever");
var globby = require("globby");
var gulp = require("gulp");
var merge = require("merge-stream");
var plumber = require("gulp-plumber");
var qunit = require("node-qunit-phantomjs");
var runSequence = require("run-sequence");
var source = require("vinyl-source-stream");
var ts = require("gulp-typescript");
var tslint = require("gulp-tslint");
var tsd = require("gulp-tsd");

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
	return gulp.src([PATHS.SRCROOT + "**/*.+(ts|tsx)", "!**/*.d.ts"])
		.pipe(ts({
			module: "commonjs"
		}))
		.pipe(gulp.dest(PATHS.BUILDROOT));
});

////////////////////////////////////////
// TSLINT
////////////////////////////////////////
//The actual task to run
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
gulp.task("bundleApi", function () {
	return browserify(PATHS.BUILDROOT + "scripts/oneNoteApi.js", { standalone: "OneNoteApi" })
		.bundle()
		.pipe(source("oneNoteApi.js"))
		.pipe(gulp.dest(PATHS.BUNDLEROOT));
});

gulp.task("bundleSampleTsModule", function () {
	return browserify(PATHS.BUILDROOT + "sample/typescript_module/sample.js")
		.bundle()
		.pipe(source("sample.js"))
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
		"bundleSampleTsModule",
		"bundleTests",
		callback);
});


////////////////////////////////////////
// EXPORT
////////////////////////////////////////
gulp.task("exportApi", function () {
	var modulesTask = gulp.src(PATHS.BUILDROOT + "scripts/**/*.js", { base: PATHS.BUILDROOT + "/scripts" })
		.pipe(gulp.dest(PATHS.TARGETROOT + "modules/"));

	var copyTask = gulp.src([
			PATHS.BUNDLEROOT + "oneNoteApi.js",
			PATHS.SRCROOT + "oneNoteApi.d.ts"
	])
		.pipe(gulp.dest(PATHS.TARGETROOT));

	return merge(modulesTask, copyTask);
});

gulp.task("exportTests", function () {
	var targetDir = PATHS.TARGETROOT + "tests/";

	var libsTask = gulp.src([
			"node_modules/qunitjs/qunit/qunit.+(css|js)",
			PATHS.SRCROOT + "tests/bind_polyfill.js"
	])
		.pipe(gulp.dest(targetDir + "libs"));

	var testsTask = gulp.src(PATHS.BUNDLEROOT + "tests/**", { base: PATHS.BUNDLEROOT + "tests/" })
		.pipe(gulp.dest(targetDir));

	var indexTask = gulp.src(PATHS.SRCROOT + "tests/index.html")
		.pipe(gulp.dest(targetDir));

	return merge(libsTask, testsTask, indexTask);
});

gulp.task("exportSampleJs", function () {
	return gulp.src([
			PATHS.BUNDLEROOT + "oneNoteApi.js",
			PATHS.SRCROOT + "sample/javascript/index.html",
			PATHS.SRCROOT + "sample/javascript/sample.js"
	])
		.pipe(gulp.dest(PATHS.WEBROOT));
});

gulp.task("exportSampleTs", function () {
	return gulp.src([
			PATHS.BUNDLEROOT + "oneNoteApi.js",
			PATHS.SRCROOT + "sample/typescript/index.html",
			PATHS.BUILDROOT + "sample/typescript/sample.js"
	])
		.pipe(gulp.dest(PATHS.WEBROOT + "typescript/"));
});

gulp.task("exportSampleTsModule", function () {
	return gulp.src([
			PATHS.SRCROOT + "sample/typescript_module/index.html",
			PATHS.BUNDLEROOT + "sample.js"
	])
		.pipe(gulp.dest(PATHS.WEBROOT + "typescript_module/"));
});

gulp.task("export", function (callback) {
	runSequence(
		"exportApi",
		"exportTests",
		"exportSampleJs",
		"exportSampleTs",
		"exportSampleTsModule",
		callback);
});

////////////////////////////////////////
// RUN
////////////////////////////////////////
gulp.task("runTests", function () {
	return qunit(PATHS.TARGETROOT + "tests/index.html");
});

////////////////////////////////////////
// SERVER
////////////////////////////////////////
gulp.task("start", function () {
	forever.startDaemon("server.js");
});

gulp.task("stop", function () {
	forever.stopAll();
});

////////////////////////////////////////
// SHORTCUT TASKS
////////////////////////////////////////
gulp.task("build", function (callback) {
	runSequence(
		"compile",
		"bundle",
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

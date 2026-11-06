# Contributing to officer

Thank you for taking an interest in officer. This document explains how the
project is run, what kinds of contributions are useful, and what to expect
after you open something.

Please read it before opening an issue or a discussion. It exists so that
your time and mine are both spent well.

## How this project is maintained

officer is maintained by one person, unpaid, alongside client work. There is
no team behind it and no service-level agreement.

Practically, this means:

- Response times vary from days to months. Silence is not a rejection, and it
  is not an approval either.
- The number of open issues I can hold in my head at once is small. Opening
  several requests in parallel does not make them move faster; it makes all of
  them move slower.
- Changes that have been discussed and agreed on are the ones that get done.

## Where to put what

| You want to…                              | Go to                                                    |
| ----------------------------------------- | -------------------------------------------------------- |
| Ask how to do something with officer      | [officeverse][book] and the [documentation][docs], then [Stack Overflow][so] with the `officer` tag or [Discussions][disc] |
| Report a reproducible bug                 | [Bug report form][bug]                                    |
| Suggest a feature or an API change        | [Discussions][disc], not the issue tracker                |

Answers to usage questions come from the community, when someone has the time.

[book]: https://ardata-fr.github.io/officeverse/
[docs]: https://davidgohel.github.io/officer/
[so]: https://stackoverflow.com/questions/tagged/officer
[bug]: https://github.com/davidgohel/officer/issues/new/choose
[disc]: https://github.com/davidgohel/officer/discussions

Issues opened in the wrong place will be moved or closed with a pointer here.
That is housekeeping, not hostility.

## Bug reports

Use the bug report form. It asks for a
[reprex](https://reprex.tidyverse.org/) and `sessionInfo()` output, and both
are required.

A reprex is minimal, self-contained, and runs in a fresh R session. Use
`tempfile()` rather than local paths. If the bug involves a specific `.docx`
or `.pptx` file, attach a stripped-down version that still reproduces it.

Reports without an executable reprex are closed without further discussion,
however carefully they are written. This is not a judgement on the report; it
is that diagnosing a rendering problem without the file that produces it is
not possible.

## Feature requests and API changes

Feature requests go to Discussions first. If a discussion converges on
something worth doing, I will open an issue for it myself.

Please describe **the problem you ran into**, not the API you would like to
see. A concrete blocked task tells me what needs to be true; a proposed
function signature only tells me one of the many ways to get there, and
usually the one that is cheapest to imagine rather than cheapest to maintain.

A good request contains:

- the task you were trying to accomplish
- the code you wrote and where it stopped working
- what you did instead, if there was a workaround
- why the workaround is not good enough

One request per discussion. A single thread proposing six changes cannot be
answered: some parts will be reasonable and some will not, and there is no
way to say yes or no to it as a whole.

### What is unlikely to be accepted

- **Changes to existing defaults or signatures.** officer's public API is
  extended downstream by flextable, rvg and mschart, and used by an
  unknown number of internal report pipelines. A change that looks like a
  small consistency fix here can break code I will never see. Deprecation
  cycles are long and expensive, so the bar for starting one is high.
- **Broad new API surface.** Every exported function is a permanent
  commitment: documented, tested, supported, and impossible to remove
  cheaply. Requests that amount to a new subsystem will generally be
  declined, not because they are bad ideas but because I cannot maintain
  them.
- **Features aimed at a tool rather than a user.** "Package X or agent Y
  would find this convenient" is not by itself a reason to change officer.
  The question is whether people are blocked on a real task.

Being declined is not a verdict on the idea. Very often the right home for it
is your own package built on top of officer.

## Pull requests

Pull requests are disabled on this repository. Changes are made by the
maintainer.

This is not a judgement on anyone's work. A pull request is a request to
review, merge, and maintain code forever, and I cannot take that on.

If you want a change, describe the problem in a discussion, as explained above,
before writing any code. Code you have already written can live in your own
fork or package.

## Code and AI

Please do not send code, whether written with an assistant or not. Thanks to AI,
or because of it, redoing a change from the idea is now much easier for me than
reviewing code written by someone else. What is valuable is the real problem you
ran into.

For the same reason, please do not open generated audits or lists of what the
package is missing. They take minutes to produce and hours to answer. A bug report
needs a reprex that you ran yourself.

## Working on the code

This is useful if you work on your own fork, or want to look at a bug yourself:

```r
install.packages(c("devtools", "testthat", "roxygen2"))
devtools::load_all()
devtools::test()
```

Conventions:

- **Documentation** is written with roxygen2, in markdown, in English. Run
  `devtools::document()` to generate the `.Rd` files.
- **Tests** use testthat (3rd edition). They check the content of the documents,
  mostly their XML, and do not need Office.

## Being civil

Be civil. Assume the other person is busy rather than rude, in both
directions. Repeated pressure on unanswered threads, or hostility toward
contributors or maintainers, will get you blocked from the repository.

## A closing note

Most of the rules above exist because of things that went wrong, including
things I got wrong, like leaving well-intentioned proposals unanswered for
far too long. The point of writing them down is that the process should not
have to be renegotiated from scratch every time, and that nobody should have
to guess where they stand.

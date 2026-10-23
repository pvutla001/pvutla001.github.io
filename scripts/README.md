# Publishing articles and scripts

Site: https://pvutla001.github.io/

## Publish an article directly in GitHub

1. Open the repository and choose **Add file → Create new file**.
2. Name the file `_posts/YYYY-MM-DD-short-title.md`, using today's date or an earlier date.
3. Start with this front matter, then write the article in Markdown:

```yaml
---
layout: post
title: "Your article title"
categories: [powershell, sharepoint]
---

A short introduction.

<!--more-->

## The problem

Describe the problem and intended outcome.

## Prerequisites

List supported versions, modules, authentication and permissions.

## Usage

Show parameters and examples.

## Verification and limitations

Explain expected output, how to verify it, and any known limitations.
```

4. Commit to `main`. GitHub Pages builds and publishes it automatically.
5. Check the Actions tab for a successful deployment. The homepage lists posts automatically.

Use fenced code blocks with `powershell` after the opening triple backticks for PowerShell examples.

## Upload a script

1. Open this `scripts` folder.
2. Choose **Add file → Upload files**, then upload the `.ps1` file.
3. Commit to `main`.
4. Link to it from a post using `[Download the script](/scripts/Your-Script.ps1)`.
5. Add prerequisites, examples, expected output, and any risks or limitations to its companion article.

For images, create an `assets/images` folder and link with `![Description](/assets/images/example.png)`.

## Edit the introduction

Open `index.html` in the repository, select Edit, make your changes, and commit.

## Work with ChatGPT or Codex

Request: "In pvutla001/pvutla001.github.io, create a post about [topic] using the post layout, add the script under scripts, and verify the Pages deployment."

The GitHub integration must have write access to this repository. If it reports access denied, enable the repository in the GitHub app's repository access settings, or use your signed-in browser.

## Publication checks

- Publish only content you have permission to share.
- Remove credentials, real tenant IDs, personal data, and internal organization details.
- Use example values for URLs, accounts and resource names.
- Test scripts before describing them as tested.
- This repository and its files are public.

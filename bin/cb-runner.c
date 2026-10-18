/* cb-runner: CalendarBridge LaunchAgent entry point. Spawns `/bin/bash -lc ~/calendarBridge/full_sync.sh`
   as a child and waits, so this binary remains the responsible process for Full Disk Access. */
#include <spawn.h>
#include <stdio.h>
#include <stdlib.h>
#include <sys/wait.h>
extern char **environ;
int main(void) {
  const char *home = getenv("HOME"); if (!home) return 70;
  char script[1024]; snprintf(script, sizeof script, "%s/calendarBridge/full_sync.sh", home);
  char *argv[] = { "/bin/bash", "-lc", script, NULL };
  pid_t pid; int rc = posix_spawn(&pid, "/bin/bash", NULL, NULL, argv, environ);
  if (rc != 0) { perror("cb-runner: posix_spawn"); return 71; }
  int st; if (waitpid(pid, &st, 0) < 0) return 72;
  return WIFEXITED(st) ? WEXITSTATUS(st) : 73;
}

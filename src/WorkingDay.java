import java.time.Duration;
import java.time.OffsetDateTime;

public class WorkingDay {
    OffsetDateTime start;
    OffsetDateTime finish;
    Duration duration;

    public OffsetDateTime getStart() {
        return start;
    }

    public void setStart() {
        this.start = OffsetDateTime.now();
    }

    public OffsetDateTime getFinish() {
        return finish;
    }

    public void setFinish() {
        this.finish = OffsetDateTime.now();
    }

    public Duration getDuration() {
        return duration;
    }

    public void calculateDuration() {
        this.duration = Duration.between(start, finish);
    }
}
